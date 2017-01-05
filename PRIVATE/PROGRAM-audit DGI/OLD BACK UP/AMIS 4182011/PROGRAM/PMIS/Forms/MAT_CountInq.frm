VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISMAT_CounterInquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Counter Inquiry"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_CountInq.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   10635
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   2580
      ScaleHeight     =   870
      ScaleWidth      =   4710
      TabIndex        =   78
      Top             =   5460
      Width           =   4710
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
         Left            =   3900
         MouseIcon       =   "MAT_CountInq.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "MAT_CountInq.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Exit Window"
         Top             =   0
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
         Left            =   3180
         MouseIcon       =   "MAT_CountInq.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "MAT_CountInq.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Move to Last Record"
         Top             =   0
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
         Left            =   2460
         MouseIcon       =   "MAT_CountInq.frx":1224
         MousePointer    =   99  'Custom
         Picture         =   "MAT_CountInq.frx":1376
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Move to First Record"
         Top             =   0
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
         Left            =   1740
         MouseIcon       =   "MAT_CountInq.frx":16D4
         MousePointer    =   99  'Custom
         Picture         =   "MAT_CountInq.frx":1826
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Find a Record"
         Top             =   0
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
         Left            =   1020
         MouseIcon       =   "MAT_CountInq.frx":1B20
         MousePointer    =   99  'Custom
         Picture         =   "MAT_CountInq.frx":1C72
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Move to Next Record"
         Top             =   0
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
         Left            =   300
         MouseIcon       =   "MAT_CountInq.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "MAT_CountInq.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   10080
      Top             =   150
   End
   Begin VB.PictureBox Frame1 
      BorderStyle     =   0  'None
      Height          =   5205
      Left            =   2700
      ScaleHeight     =   5205
      ScaleWidth      =   7875
      TabIndex        =   15
      Top             =   120
      Width           =   7875
      Begin VB.TextBox txtSSL 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   6960
         TabIndex        =   71
         Text            =   "Text1"
         Top             =   4770
         Width           =   825
      End
      Begin VB.TextBox txtLast_RecD 
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
         Height          =   345
         Left            =   6510
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   2340
         Width           =   1275
      End
      Begin VB.TextBox txtLast_RecQ 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   6510
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   1950
         Width           =   1275
      End
      Begin VB.TextBox txtLastM_Sell 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   6510
         TabIndex        =   67
         Text            =   "Text1"
         Top             =   1230
         Width           =   1275
      End
      Begin VB.TextBox txtLastM_MAD 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   6510
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox txtMAD 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   3930
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txtPromo_SRP 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   3930
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   570
         Width           =   915
      End
      Begin VB.TextBox txtPromo_WFP 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   3930
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   -540
         Width           =   915
      End
      Begin VB.TextBox txtSRP 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   3930
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   -990
         Width           =   915
      End
      Begin VB.TextBox txtWFP 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   3930
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   -600
         Width           =   915
      End
      Begin VB.TextBox txtNOShip 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   4770
         Width           =   855
      End
      Begin VB.TextBox txtResService 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   4140
         Width           =   855
      End
      Begin VB.TextBox txtCompOnHand 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   3390
         Width           =   855
      End
      Begin VB.TextBox txtIssuances 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   2970
         Width           =   855
      End
      Begin VB.TextBox txtReceipts 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   2550
         Width           =   855
      End
      Begin VB.TextBox txtTissqty 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtTrecQty 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   1770
         Width           =   855
      End
      Begin VB.TextBox txtTPOQty 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox txtOnORder 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   930
         Width           =   855
      End
      Begin VB.TextBox txtOnHand 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   540
         Width           =   855
      End
      Begin VB.TextBox txtPartDesc 
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
         Height          =   345
         Left            =   4500
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   60
         Width           =   3285
      End
      Begin VB.TextBox txtPartNo 
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
         Height          =   345
         Left            =   1500
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   60
         Width           =   1815
      End
      Begin VB.TextBox txtPriceClass 
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
         Height          =   345
         Left            =   3930
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1440
         Width           =   435
      End
      Begin VB.TextBox txtINVClass 
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
         Height          =   345
         Left            =   3930
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1830
         Width           =   435
      End
      Begin VB.TextBox txtSubInvClas 
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
         Height          =   345
         Left            =   3930
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2220
         Width           =   435
      End
      Begin VB.TextBox txtVehType 
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
         Height          =   345
         Left            =   3930
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2610
         Width           =   435
      End
      Begin VB.TextBox txtLocation 
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
         Height          =   705
         Left            =   3270
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3030
         Width           =   1605
      End
      Begin VB.TextBox txtModelCode 
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
         Height          =   855
         Left            =   2340
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   4260
         Width           =   2535
      End
      Begin VB.TextBox txtOldNo 
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
         Height          =   345
         Left            =   6510
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2730
         Width           =   1275
      End
      Begin VB.TextBox txtNewNo 
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
         Height          =   345
         Left            =   6510
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3120
         Width           =   1275
      End
      Begin VB.TextBox txtGenNo 
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
         Height          =   345
         Left            =   6510
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3510
         Width           =   1275
      End
      Begin VB.TextBox txtSupCode 
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
         Height          =   345
         Left            =   6510
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   3900
         Width           =   1275
      End
      Begin VB.TextBox txtDateEntered 
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
         Height          =   345
         Left            =   6510
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   4290
         Width           =   1275
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Safety Stock Level"
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
         Height          =   225
         Left            =   5070
         TabIndex        =   70
         Top             =   4830
         Width           =   2895
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SRP"
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
         Left            =   2340
         TabIndex        =   49
         Top             =   -960
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   3360
         TabIndex        =   48
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Material Code"
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
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "LM On Hand"
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
         Left            =   120
         TabIndex        =   46
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "On-Order"
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
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Rec Qty"
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
         Left            =   120
         TabIndex        =   44
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PO Qty"
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
         Left            =   120
         TabIndex        =   43
         Top             =   1410
         Width           =   1155
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt"
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
         Left            =   120
         TabIndex        =   42
         Top             =   2580
         Width           =   1155
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Issuance"
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
         Left            =   120
         TabIndex        =   41
         Top             =   2970
         Width           =   1155
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Computed On Hand"
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
         Height          =   435
         Left            =   120
         TabIndex        =   40
         Top             =   3330
         Width           =   1155
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reserved for Service"
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
         Height          =   465
         Left            =   120
         TabIndex        =   39
         Top             =   3900
         Width           =   1665
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Shipping Months"
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
         Height          =   435
         Left            =   120
         TabIndex        =   38
         Top             =   4530
         Width           =   1665
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PAMCOR WFP"
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
         Left            =   2340
         TabIndex        =   37
         Top             =   -570
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PROMO WFP"
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
         Left            =   2340
         TabIndex        =   36
         Top             =   -510
         Width           =   1575
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "S.R.P."
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
         Left            =   2340
         TabIndex        =   35
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Moving Ave. On Demand"
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
         Height          =   405
         Left            =   2340
         TabIndex        =   34
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label21 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Price Class"
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
         Left            =   2340
         TabIndex        =   33
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Class"
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
         Height          =   225
         Left            =   2340
         TabIndex        =   32
         Top             =   1890
         Width           =   1575
      End
      Begin VB.Label Label23 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Class"
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
         Left            =   2340
         TabIndex        =   31
         Top             =   2250
         Width           =   1575
      End
      Begin VB.Label Label24 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Type"
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
         Left            =   2340
         TabIndex        =   30
         Top             =   2670
         Width           =   1575
      End
      Begin VB.Label Label25 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Left            =   2340
         TabIndex        =   29
         Top             =   3060
         Width           =   1575
      End
      Begin VB.Label Label26 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
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
         Left            =   2370
         TabIndex        =   28
         Top             =   3990
         Width           =   1575
      End
      Begin VB.Label Label27 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "LAST MONTH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5070
         TabIndex        =   27
         Top             =   540
         Width           =   2295
      End
      Begin VB.Label Label29 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "M.A.D."
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
         Left            =   5070
         TabIndex        =   26
         Top             =   870
         Width           =   1605
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sell Price"
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
         Left            =   5070
         TabIndex        =   25
         Top             =   1290
         Width           =   1605
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "LAST RECEIVED (RR)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5070
         TabIndex        =   24
         Top             =   1650
         Width           =   2295
      End
      Begin VB.Label Label32 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   5070
         TabIndex        =   23
         Top             =   1980
         Width           =   1605
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   5070
         TabIndex        =   22
         Top             =   2370
         Width           =   1605
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Old Number"
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
         Left            =   5070
         TabIndex        =   21
         Top             =   2760
         Width           =   1605
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "New Number"
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
         Left            =   5070
         TabIndex        =   20
         Top             =   3150
         Width           =   1605
      End
      Begin VB.Label Label35 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Generic No."
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
         Left            =   5070
         TabIndex        =   19
         Top             =   3540
         Width           =   1605
      End
      Begin VB.Label Label36 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
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
         Left            =   5070
         TabIndex        =   18
         Top             =   3930
         Width           =   1605
      End
      Begin VB.Label Label37 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Entered"
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
         Left            =   5070
         TabIndex        =   17
         Top             =   4350
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ISS Qty"
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
         Left            =   120
         TabIndex        =   16
         Top             =   2190
         Width           =   1155
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   6255
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   2595
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
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   1050
         Width           =   2475
      End
      Begin VB.OptionButton optDescription 
         Caption         =   "D&escription [Alt + E]"
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
         Left            =   180
         TabIndex        =   74
         Top             =   720
         Width           =   2385
      End
      Begin VB.OptionButton optPartNo 
         Caption         =   "&Material Code [Alt + M]"
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
         Left            =   180
         TabIndex        =   73
         Top             =   480
         Value           =   -1  'True
         Width           =   2385
      End
      Begin MSComctlLib.ListView lstParts 
         Height          =   4755
         Left            =   30
         TabIndex        =   76
         Top             =   1440
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   8387
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
         MouseIcon       =   "MAT_CountInq.frx":247B
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PART NUMBER"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label15 
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
         TabIndex        =   77
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.Label labSSL 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "BELOW SAFETY STOCK LEVEL!"
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
      Height          =   615
      Left            =   7410
      TabIndex        =   50
      Top             =   5640
      Width           =   3285
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   3390
      TabIndex        =   14
      Top             =   420
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   3720
      TabIndex        =   13
      Top             =   270
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmPMISMAT_CounterInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPartMas                                                         As ADODB.Recordset

Function PartFound(ByVal str2find) As Boolean
    On Error GoTo BFoundErr
    Dim result                                                        As Boolean
    Dim rsBClone                                                      As ADODB.Recordset
    result = False
    If Not IsNull(str2find) Then
        Set rsBClone = New ADODB.Recordset
        Set rsBClone = rsPartMas.Clone

        rsBClone.Find "STOCKNO = '" & str2find & "'"
        result = Not rsBClone.EOF
        If result Then
            rsPartMas.Bookmark = rsBClone.Bookmark
        End If
        Set rsBClone = Nothing
    End If
    PartFound = result
    Exit Function
BFoundErr:
    ShowVBError
End Function

Function PartFound2(ByVal str2find) As Boolean
    On Error GoTo BFoundErr
    Dim result                                                        As Boolean
    Dim rsBClone                                                      As ADODB.Recordset
    result = False
    If Not IsNull(str2find) Then
        Set rsBClone = New ADODB.Recordset
        Set rsBClone = rsPartMas.Clone

        rsBClone.Find "STOCKDESC = '" & str2find & "'"
        result = Not rsBClone.EOF
        If result Then
            rsPartMas.Bookmark = rsBClone.Bookmark
        End If
        Set rsBClone = Nothing
    End If
    PartFound2 = result
    Exit Function
BFoundErr:
    ShowVBError
End Function

Sub initMemvars()
    txtPartNo.Text = ""
    txtPartDesc.Text = ""
    txtOnHand.Text = ""
    txtOnORder.Text = ""
    txtTPOQty.Text = ""
    txtTrecQty.Text = ""
    txtTISSQty.Text = ""
    txtReceipts.Text = ""
    txtIssuances.Text = ""
    txtCompOnHand.Text = ""
    txtResService.Text = ""
    txtNOShip.Text = ""
    txtWFP.Text = ""
    txtSRP.Text = ""
    txtPromo_WFP.Text = ""
    txtPromo_SRP.Text = ""
    txtMAD.Text = ""
    txtPriceClass.Text = ""
    txtINVClass.Text = ""
    txtSubInvClas.Text = ""
    txtVehType.Text = ""
    txtLocation.Text = ""
    txtModelCode.Text = ""
    txtLastM_MAD.Text = ""
    txtLastM_Sell.Text = ""
    txtLast_RecQ.Text = ""
    txtLast_RecD.Text = ""
    txtOldNo.Text = ""
    txtNewNo.Text = ""
    txtGenNo.Text = ""
    txtSupCode.Text = ""
    txtDateEntered.Text = ""
    txtSSL.Text = ""
End Sub

Sub StoreMemvars()
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        labID.Caption = rsPartMas!ID
        txtPartNo.Text = Null2String(rsPartMas!STOCKNO)
        txtPartDesc.Text = Null2String(rsPartMas!STOCKDESC)
        txtOnHand.Text = N2Str2IntZero(rsPartMas!lastm_oh)
        txtOnORder.Text = N2Str2IntZero(rsPartMas!onorder)
        txtTPOQty.Text = N2Str2IntZero(rsPartMas!tpoqty)
        txtTrecQty.Text = N2Str2IntZero(rsPartMas!trecqty)
        txtTISSQty.Text = N2Str2IntZero(rsPartMas!tissqty)
        txtReceipts.Text = N2Str2IntZero(rsPartMas!receipts)
        txtIssuances.Text = N2Str2IntZero(rsPartMas!issuances)
        txtCompOnHand.Text = N2Str2IntZero(rsPartMas!ONHAND)
        txtResService.Text = N2Str2IntZero(rsPartMas!RESSERVICE)
        txtNOShip.Text = N2Str2IntZero(rsPartMas!NOSHIP)
        txtWFP.Text = N2Str2Zero(rsPartMas!WFP)
        txtSRP.Text = N2Str2Zero(rsPartMas!SRP)
        txtPromo_WFP.Text = N2Str2Zero(rsPartMas!promo_wfp)
        txtPromo_SRP.Text = N2Str2Zero(rsPartMas!promo_srp)
        txtMAD.Text = N2Str2IntZero(rsPartMas!mad)
        txtPriceClass.Text = Null2String(rsPartMas!priceclass)
        txtINVClass.Text = Null2String(rsPartMas!InvClass)
        txtSubInvClas.Text = Null2String(rsPartMas!SubInvClas)
        txtVehType.Text = Null2String(rsPartMas!vehtype)
        txtLocation.Text = Null2String(rsPartMas!Location)
        txtModelCode.Text = Null2String(rsPartMas!modelcode)
        txtLastM_MAD.Text = N2Str2IntZero(rsPartMas!LASTM_MAD)
        txtLastM_Sell.Text = N2Str2Zero(rsPartMas!LASTM_SELL)
        txtLast_RecQ.Text = Null2String(rsPartMas!last_recq)
        txtLast_RecD.Text = Null2String(rsPartMas!LAST_RECD)
        txtOldNo.Text = Null2String(rsPartMas!oldno)
        txtNewNo.Text = Null2String(rsPartMas!NEWNO)
        txtGenNo.Text = Null2String(rsPartMas!GENNO)
        txtSupCode.Text = Null2String(rsPartMas!SupCode)
        txtDateEntered.Text = Null2String(rsPartMas!DATE_ENTERED)
        txtSSL.Text = Null2String(rsPartMas!SSTOCK)
    Else
        ShowNoRecord
    End If
End Sub

Sub rsRefresh()
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "select * from CSMS_MATMAS WHERE [TYPE] = 'M' order by STOCKNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub FillGrid()
    Dim rsParts                                                       As ADODB.Recordset
    lstParts.Enabled = False
    lstParts.Sorted = False: lstParts.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    Set rsParts = gconDMIS.Execute("select STOCKNO,STOCKNO x from CSMS_MATMAS order by STOCKNO asc")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstParts.Enabled = True: Listview_Loadval Me.lstParts.ListItems, rsParts: lstParts.Refresh
    Else
        lstParts.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsParts                                                       As ADODB.Recordset
    lstParts.Enabled = False
    lstParts.Sorted = False: lstParts.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsParts = gconDMIS.Execute("select STOCKNO, STOCKNO from CSMS_MATMAS where STOCKNO like'" & XXX & "%'")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstParts.Enabled = True: Listview_Loadval Me.lstParts.ListItems, rsParts: lstParts.Refresh
    Else
        lstParts.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim rsParts                                                       As ADODB.Recordset
    lstParts.Enabled = False
    lstParts.Sorted = False: lstParts.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    Set rsParts = gconDMIS.Execute("select STOCKDESC, STOCKNO from CSMS_MATMAS order by STOCKDESC asc")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstParts.Enabled = True: Listview_Loadval Me.lstParts.ListItems, rsParts: lstParts.Refresh
    Else
        lstParts.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsParts                                                       As ADODB.Recordset
    lstParts.Enabled = False
    lstParts.Sorted = False: lstParts.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsParts = gconDMIS.Execute("select STOCKDESC, STOCKNO from CSMS_MATMAS where STOCKDESC like'" & XXX & "%'")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstParts.Enabled = True: Listview_Loadval Me.lstParts.ListItems, rsParts: lstParts.Refresh
    Else
        lstParts.Enabled = False
    End If
End Sub

Private Sub cmdFirst_Click()
    rsPartMas.MoveFirst
    StoreMemvars
End Sub

Private Sub cmdLast_Click()
    rsPartMas.MoveLast
    StoreMemvars
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    textSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsPartMas.MoveNext
    If rsPartMas.EOF Then
        rsPartMas.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    rsPartMas.MovePrevious
    If rsPartMas.BOF Then
        rsPartMas.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    Frame1.Enabled = False
    textSearch.Text = ""
    initMemvars
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub Timer1_Timer()
    If NumericVal(txtCompOnHand.Text) < NumericVal(txtSSL.Text) Then
        labSSL.Caption = "Below Safety Stock Level"
        If labSSL.Visible = False Then
            labSSL.Visible = True
        Else
            labSSL.Visible = False
        End If
    Else
        labSSL.Caption = ""
    End If
End Sub

Private Sub lstParts_GotFocus()
    rsPartMas.Bookmark = rsFind(rsPartMas.Clone, "STOCKNO", lstParts.SelectedItem.SubItems(1)).Bookmark
    StoreMemvars
End Sub

Private Sub lstParts_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsPartMas.Bookmark = rsFind(rsPartMas.Clone, "STOCKNO", lstParts.SelectedItem.SubItems(1)).Bookmark
    StoreMemvars
End Sub

Private Sub lstParts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstParts
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstParts_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then On Error Resume Next: textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If optPartNo.Value = True Then
        If Trim(textSearch.Text) = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    Else
        If Trim(textSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        On Error Resume Next
    End If
End Sub

Private Sub optDescription_Click()
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optPARTNO_Click()
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

