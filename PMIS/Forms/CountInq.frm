VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISInquiry_CounterInquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Counter Inquiry"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   ForeColor       =   &H00DEDFDE&
   Icon            =   "CountInq.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   11010
   Begin VB.OptionButton optPartNo 
      Caption         =   "Pa&rt Number [Alt + R]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   71
      Top             =   270
      Value           =   -1  'True
      Width           =   2385
   End
   Begin VB.OptionButton optDescription 
      Caption         =   "D&escription [Alt + E]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   70
      Top             =   480
      Width           =   2385
   End
   Begin VB.TextBox textSearch 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   810
      Width           =   2445
   End
   Begin VB.TextBox txtOnHand 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   65
      Text            =   "Text1"
      Top             =   1010
      Width           =   855
   End
   Begin VB.TextBox txtLastM_MAD 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   1010
      Width           =   1275
   End
   Begin VB.TextBox txtSellingPrice 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6780
      Locked          =   -1  'True
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   1010
      Width           =   1005
   End
   Begin VB.TextBox txtNOShip 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   4620
      Width           =   855
   End
   Begin VB.TextBox txtDateEntered 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   4218
      Width           =   1275
   End
   Begin VB.TextBox txtSupCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   3817
      Width           =   1275
   End
   Begin VB.TextBox txtGenNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   3416
      Width           =   1275
   End
   Begin VB.TextBox txtNewNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   3015
      Width           =   1275
   End
   Begin VB.TextBox txtOldNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   2614
      Width           =   1275
   End
   Begin VB.TextBox txtModelCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   5100
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   4290
      Width           =   2745
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   5100
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   3510
      Width           =   2715
   End
   Begin VB.TextBox txtVehType 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6780
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   2970
      Width           =   1035
   End
   Begin VB.TextBox txtSubInvClas 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6780
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   2565
      Width           =   1035
   End
   Begin VB.TextBox txtINVClass 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6780
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   2175
      Width           =   1035
   End
   Begin VB.TextBox txtPriceClass 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6780
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   1785
      Width           =   1035
   End
   Begin VB.TextBox txtPartNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   180
      Width           =   1995
   End
   Begin VB.TextBox txtPartDesc 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   595
      Width           =   6765
   End
   Begin VB.TextBox txtOnORder 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1410
      Width           =   855
   End
   Begin VB.TextBox txtTPOQty 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1810
      Width           =   855
   End
   Begin VB.TextBox txtTrecQty 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   2210
      Width           =   855
   End
   Begin VB.TextBox txtTissqty 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   2610
      Width           =   855
   End
   Begin VB.TextBox txtReceipts 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3010
      Width           =   855
   End
   Begin VB.TextBox txtIssuances 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3410
      Width           =   855
   End
   Begin VB.TextBox txtCompOnHand 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3810
      Width           =   855
   End
   Begin VB.TextBox txtResService 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4210
      Width           =   855
   End
   Begin VB.TextBox txtMAD 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6780
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1395
      Width           =   1035
   End
   Begin VB.TextBox txtLastM_Sell 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1411
      Width           =   1275
   End
   Begin VB.TextBox txtLast_RecQ 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1812
      Width           =   1275
   End
   Begin VB.TextBox txtLast_RecD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2213
      Width           =   1275
   End
   Begin VB.TextBox txtSSL 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4620
      Width           =   1305
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   6450
      ScaleHeight     =   930
      ScaleWidth      =   4530
      TabIndex        =   2
      Top             =   5070
      Width           =   4530
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
         Left            =   3720
         MouseIcon       =   "CountInq.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CountInq.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   3000
         MouseIcon       =   "CountInq.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "CountInq.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   2280
         MouseIcon       =   "CountInq.frx":1224
         MousePointer    =   99  'Custom
         Picture         =   "CountInq.frx":1376
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   1560
         MouseIcon       =   "CountInq.frx":16D4
         MousePointer    =   99  'Custom
         Picture         =   "CountInq.frx":1826
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   840
         MouseIcon       =   "CountInq.frx":1B20
         MousePointer    =   99  'Custom
         Picture         =   "CountInq.frx":1C72
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   120
         MouseIcon       =   "CountInq.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "CountInq.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9780
      Top             =   150
   End
   Begin MSComctlLib.ListView lstParts 
      Height          =   4665
      Left            =   30
      TabIndex        =   72
      Top             =   1230
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   8229
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "CountInq.frx":247B
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   73
      Top             =   60
      Width           =   1455
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Month M.A.D."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8115
      TabIndex        =   68
      Top             =   1080
      Width           =   1485
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "S.R.P."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   6315
      TabIndex        =   67
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "LM On Hand"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3090
      TabIndex        =   66
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Shipping Months"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   2955
      TabIndex        =   62
      Top             =   4560
      Width           =   1110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Issuance Qty"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2625
      TabIndex        =   60
      Top             =   2625
      Width           =   1440
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Entered"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8550
      TabIndex        =   59
      Top             =   4290
      Width           =   1050
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8430
      TabIndex        =   58
      Top             =   3870
      Width           =   1170
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Generic No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8655
      TabIndex        =   57
      Top             =   3480
      Width           =   945
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "New Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8535
      TabIndex        =   56
      Top             =   3030
      Width           =   1065
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8625
      TabIndex        =   55
      Top             =   2625
      Width           =   975
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Received Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8040
      TabIndex        =   54
      Top             =   2235
      Width           =   1560
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Month Received  Qty"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   7845
      TabIndex        =   53
      Top             =   1875
      Width           =   1755
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Month Sell Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   7860
      TabIndex        =   52
      Top             =   1455
      Width           =   1740
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Model Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   5100
      TabIndex        =   51
      Top             =   4080
      Width           =   1140
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5100
      TabIndex        =   50
      Top             =   3300
      Width           =   855
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5700
      TabIndex        =   49
      Top             =   3030
      Width           =   1065
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Sub-Class"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5925
      TabIndex        =   48
      Top             =   2625
      Width           =   840
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory Class"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5475
      TabIndex        =   47
      Top             =   2235
      Width           =   1290
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Price Class"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5835
      TabIndex        =   46
      Top             =   1845
      Width           =   930
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Moving Ave. On Demand"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   5040
      TabIndex        =   45
      Top             =   1455
      Width           =   1710
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Reserved for Service"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2655
      TabIndex        =   44
      Top             =   4245
      Width           =   1410
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Computed On Hand"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2700
      TabIndex        =   43
      Top             =   3840
      Width           =   1365
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Issuance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3375
      TabIndex        =   42
      Top             =   3435
      Width           =   750
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3510
      TabIndex        =   41
      Top             =   3030
      Width           =   615
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Current PO Qty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2835
      TabIndex        =   40
      Top             =   1845
      Width           =   1230
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Received Qty"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2610
      TabIndex        =   39
      Top             =   2220
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "On-Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3360
      TabIndex        =   38
      Top             =   1455
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3090
      TabIndex        =   37
      Top             =   240
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3180
      TabIndex        =   36
      Top             =   645
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Safety Stock Level"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8085
      TabIndex        =   35
      Top             =   4680
      Width           =   1515
   End
   Begin VB.Label labSSL 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "BELOW SAFETY STOCK LEVEL!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   6450
      TabIndex        =   1
      Top             =   210
      Width           =   3285
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   8040
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "frmPMISInquiry_CounterInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSPARTMAS                                          As ADODB.Recordset
Dim LOCAL_STOCKTYPE                                    As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    RSPARTMAS.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    RSPARTMAS.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    RSPARTMAS.MoveNext
    If RSPARTMAS.EOF Then
        RSPARTMAS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    RSPARTMAS.MovePrevious
    If RSPARTMAS.BOF Then
        RSPARTMAS.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Sub FillGrid()
    Dim rsParts                                        As ADODB.Recordset
    If optPartNo.Value = True Then
        If LTrim(RTrim(textSearch)) = "" Then
            Set rsParts = gconDMIS.Execute("SELECT TOP 200  STOCKNO,ID FROM PMIS_STOCKMAS WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND ACTIVE = 'Y' ORDER BY STOCKNO ASC")
        Else
            Set rsParts = gconDMIS.Execute("SELECT STOCKNO,ID FROM PMIS_STOCKMAS WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND ACTIVE = 'Y' AND STOCKNO like " & N2Str2Null(textSearch & "%") & " ORDER BY STOCKNO ASC")
        End If
    Else
        If LTrim(RTrim(textSearch)) = "" Then
            Set rsParts = gconDMIS.Execute("SELECT TOP 200  STOCKDESC,ID FROM PMIS_STOCKMAS WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND ACTIVE = 'Y' ORDER BY STOCKDESC ASC")
        Else
            Set rsParts = gconDMIS.Execute("SELECT TOP 200  STOCKDESC,ID FROM PMIS_STOCKMAS WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND ACTIVE = 'Y' AND STOCKDESC like " & N2Str2Null(textSearch & "%") & " ORDER BY STOCKDESC ASC")
        End If
    End If
    lstParts.Sorted = False
    Listview_Loadval Me.lstParts.ListItems, rsParts
    lstParts.Sorted = True
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    textSearch.Text = ""
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISInquiry_CounterInquiry = Nothing
    UnloadForm Me
End Sub

Sub initMemvars()
    txtPartNo.Text = ""
    txtPartDesc.Text = ""
    txtOnHand.Text = ""
    txtOnORder.Text = ""
    txtTPOQty.Text = ""
    txtTrecQty.Text = ""
    txtTissqty.Text = ""
    txtReceipts.Text = ""
    txtIssuances.Text = ""
    txtCompOnHand.Text = ""
    txtResService.Text = ""
    txtNOShip.Text = ""
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

Private Sub lstParts_GotFocus()
    If lstParts.SelectedItem Is Nothing Then Exit Sub
    On Error Resume Next
    RSPARTMAS.MoveFirst
    RSPARTMAS.Find ("ID=" & lstParts.SelectedItem.ListSubItems(1).Text)
    StoreMemVars
End Sub

Private Sub lstParts_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    RSPARTMAS.MoveFirst
    RSPARTMAS.Find ("ID=" & Item.ListSubItems(1).Text)
    StoreMemVars
End Sub

Private Sub lstParts_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub optDescription_Click()
    FillGrid
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optPARTNO_Click()
    FillGrid
    On Error Resume Next
    textSearch.SetFocus
End Sub

Sub rsRefresh()
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "select * from PMIS_STOCKMAS WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' and ACTIVE = 'Y' order by STOCKNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub SETSTOCK_TYPE(XXX As String)
    LOCAL_STOCKTYPE = XXX
End Sub

Sub StoreMemVars()
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        labid.Caption = RSPARTMAS!ID
        txtPartNo.Text = Null2String(RSPARTMAS!STOCKNO)
        txtPartDesc.Text = Null2String(RSPARTMAS!STOCKDESC)
        txtOnHand.Text = N2Str2IntZero(RSPARTMAS!LASTM_OH)
        txtOnORder.Text = N2Str2IntZero(RSPARTMAS!ONORDER)
        txtTPOQty.Text = N2Str2IntZero(RSPARTMAS!tpoqty)
        txtTrecQty.Text = N2Str2IntZero(RSPARTMAS!TRECQTY)
        txtTissqty.Text = N2Str2IntZero(RSPARTMAS!TISSQTY)
        txtReceipts.Text = N2Str2IntZero(RSPARTMAS!RECEIPTS)
        txtIssuances.Text = N2Str2IntZero(RSPARTMAS!ISSUANCES)
        txtCompOnHand.Text = N2Str2IntZero(RSPARTMAS!ONHAND)
        txtResService.Text = N2Str2IntZero(RSPARTMAS!RESSERVICE)
        txtNOShip.Text = N2Str2IntZero(RSPARTMAS!NOSHIP)
        txtSellingPrice = FormatNumber(ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP)))
        txtMAD.Text = N2Str2IntZero(RSPARTMAS!mad)
        txtPriceClass.Text = Null2String(RSPARTMAS!priceclass)
        txtINVClass.Text = Null2String(RSPARTMAS!InvClass)
        txtSubInvClas.Text = Null2String(RSPARTMAS!SubInvClas)
        txtVehType.Text = Null2String(RSPARTMAS!vehtype)
        txtLocation.Text = Null2String(RSPARTMAS!Location)
        txtModelCode.Text = Null2String(RSPARTMAS!MODELCODE)
        txtLastM_MAD.Text = N2Str2IntZero(RSPARTMAS!LASTM_MAD)
        txtLastM_Sell.Text = FormatNumber(N2Str2Zero(RSPARTMAS!LASTM_SELL))
        txtLast_RecQ.Text = Null2String(RSPARTMAS!last_recq)
        txtLast_RecD.Text = Null2String(RSPARTMAS!LAST_RECD)
        txtOldNo.Text = Null2String(RSPARTMAS!oldno)
        txtNewNo.Text = Null2String(RSPARTMAS!NEWNO)
        txtGenNo.Text = Null2String(RSPARTMAS!GENNO)
        txtSupCode.Text = Null2String(RSPARTMAS!SupCode)
        txtDateEntered.Text = Null2String(RSPARTMAS!DATE_ENTERED)
        txtSSL.Text = Null2String(RSPARTMAS!SSTOCK)
        
        If N2Str2IntZero(RSPARTMAS!ONHAND) < N2Str2IntZero(RSPARTMAS!SSTOCK) Then
            labSSL.Caption = "BELOW SAFETY STOCK LEVEL!"
        Else
            labSSL.Caption = ""
        End If
        
    Else
        ShowNoRecord
    End If
End Sub

Private Sub textSearch_Change()
    FillGrid
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstParts.ListItems.Count > 0 Then
            lstParts.SetFocus
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    If txtCompOnHand.Text < txtSSL.Text Then
        labSSL.Caption = "***BELOW SAFETY STOCK LEVEL***"
        If labSSL.Visible = False Then
            labSSL.Visible = True
        Else
            labSSL.Visible = False
        End If
    Else
        labSSL.Caption = ""
    End If
End Sub

