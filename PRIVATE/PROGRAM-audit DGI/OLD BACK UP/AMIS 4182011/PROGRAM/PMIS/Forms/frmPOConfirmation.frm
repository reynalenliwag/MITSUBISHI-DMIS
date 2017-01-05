VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "WIZFLEX.OCX"
Begin VB.Form frmPMISPOConfirmation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PO Confirmation"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPOC 
      BackColor       =   &H80000015&
      Height          =   765
      Left            =   3810
      ScaleHeight     =   705
      ScaleWidth      =   4095
      TabIndex        =   46
      Top             =   3990
      Width           =   4155
      Begin wizProgBar.Prg progPOC 
         Height          =   315
         Left            =   60
         TabIndex        =   47
         Top             =   300
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         Picture         =   "frmPOConfirmation.frx":0000
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   0
         BarPicture      =   "frmPOConfirmation.frx":001C
         ShowText        =   -1  'True
         Text            =   "Saving"
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
      Begin VB.Label labPOC 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   60
         TabIndex        =   48
         Top             =   30
         Width           =   5595
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Parts Order Details"
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
      Height          =   5415
      Left            =   8010
      TabIndex        =   19
      Top             =   1290
      Width           =   3615
      Begin VB.TextBox txtModelCode 
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   40
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   330
         Width           =   1515
      End
      Begin VB.TextBox txtSegment 
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   39
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   690
         Width           =   1515
      End
      Begin VB.TextBox txtSOCategory 
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   38
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   1080
         Width           =   1515
      End
      Begin VB.TextBox txtPartsOrigin 
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   37
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   1500
         Width           =   1515
      End
      Begin VB.TextBox txtByRegion 
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   36
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   1860
         Width           =   1515
      End
      Begin VB.TextBox txtBackOrderAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   34
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   4920
         Width           =   1515
      End
      Begin VB.TextBox txtUnservedAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   32
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   4560
         Width           =   1515
      End
      Begin VB.TextBox txtAllocAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   30
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   4200
         Width           =   1515
      End
      Begin VB.TextBox txtOrderAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   28
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   3840
         Width           =   1515
      End
      Begin VB.TextBox txtUnitPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   26
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   3480
         Width           =   1515
      End
      Begin VB.TextBox txtQty_BackOrder 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   24
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   3060
         Width           =   1515
      End
      Begin VB.TextBox txtQty_FillRate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   22
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   2700
         Width           =   1515
      End
      Begin VB.TextBox txtQty_Unserved 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   1950
         TabIndex        =   20
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   2340
         Width           =   1515
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
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
         Left            =   330
         TabIndex        =   45
         Top             =   390
         Width           =   1845
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Segment"
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
         Left            =   330
         TabIndex        =   44
         Top             =   750
         Width           =   1845
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "SO Category"
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
         Left            =   330
         TabIndex        =   43
         Top             =   1140
         Width           =   1845
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Parts Origin"
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
         Left            =   330
         TabIndex        =   42
         Top             =   1560
         Width           =   1845
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Parts Region"
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
         Left            =   330
         TabIndex        =   41
         Top             =   1920
         Width           =   1845
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Back Order Amt."
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
         Left            =   330
         TabIndex        =   35
         Top             =   4980
         Width           =   1845
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Unserved Amt."
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
         Left            =   330
         TabIndex        =   33
         Top             =   4620
         Width           =   1845
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Allocated Amt."
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
         Left            =   330
         TabIndex        =   31
         Top             =   4260
         Width           =   1845
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ordered Amt."
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
         Left            =   330
         TabIndex        =   29
         Top             =   3900
         Width           =   1845
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
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
         Left            =   330
         TabIndex        =   27
         Top             =   3540
         Width           =   1845
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Back Order"
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
         Left            =   330
         TabIndex        =   25
         Top             =   3120
         Width           =   1845
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Fill Rate"
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
         Left            =   330
         TabIndex        =   23
         Top             =   2760
         Width           =   1845
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Unserved"
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
         Left            =   330
         TabIndex        =   21
         Top             =   2400
         Width           =   1845
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PO Details"
      Enabled         =   0   'False
      Height          =   2385
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   6885
      Begin VB.TextBox txtSONum 
         BackColor       =   &H00FFC0C0&
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
         Left            =   1560
         TabIndex        =   17
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   1830
         Width           =   1515
      End
      Begin VB.TextBox txtSOYear 
         BackColor       =   &H00FFC0C0&
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
         Left            =   1560
         TabIndex        =   15
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   1470
         Width           =   1515
      End
      Begin VB.TextBox txtSOType 
         BackColor       =   &H00FFC0C0&
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
         Left            =   1560
         TabIndex        =   13
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   750
         Width           =   1515
      End
      Begin VB.TextBox txtPODate 
         BackColor       =   &H00FFC0C0&
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
         Left            =   5340
         MaxLength       =   10
         TabIndex        =   11
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   1080
         Width           =   1395
      End
      Begin VB.TextBox txtPONO 
         BackColor       =   &H00FFC0C0&
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
         Left            =   5340
         MaxLength       =   10
         TabIndex        =   9
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   690
         Width           =   1395
      End
      Begin VB.TextBox txtDealerName 
         BackColor       =   &H00FFC0C0&
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
         Left            =   2130
         TabIndex        =   8
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   240
         Width           =   4605
      End
      Begin VB.TextBox txtSOMonth 
         BackColor       =   &H00FFC0C0&
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
         Left            =   1560
         TabIndex        =   5
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   1110
         Width           =   1515
      End
      Begin VB.TextBox txtDealerCode 
         BackColor       =   &H00FFC0C0&
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
         Left            =   1050
         MaxLength       =   6
         TabIndex        =   4
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   240
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         Height          =   1605
         Left            =   300
         Top             =   660
         Width           =   2865
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "SO Number"
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
         Left            =   390
         TabIndex        =   18
         Top             =   1860
         Width           =   1845
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "SO Year"
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
         Left            =   390
         TabIndex        =   16
         Top             =   1500
         Width           =   1845
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SO Type"
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
         Left            =   390
         TabIndex        =   14
         Top             =   810
         Width           =   1845
      End
      Begin VB.Label Label2 
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
         Height          =   315
         Left            =   4110
         TabIndex        =   12
         Top             =   1140
         Width           =   1845
      End
      Begin VB.Label Label1 
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
         Height          =   315
         Left            =   4110
         TabIndex        =   10
         Top             =   750
         Width           =   1845
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "SO Month"
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
         Left            =   390
         TabIndex        =   7
         Top             =   1140
         Width           =   1845
      End
      Begin VB.Label Label21 
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
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Top             =   300
         Width           =   1845
      End
   End
   Begin VB.CommandButton cmdSavePOConfirmation 
      Caption         =   "&Save / Confirmed Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   8040
      MaskColor       =   &H0000FFFF&
      Picture         =   "frmPOConfirmation.frx":0038
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6810
      Width           =   1965
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
      Height          =   825
      Left            =   10710
      MaskColor       =   &H0000FFFF&
      Picture         =   "frmPOConfirmation.frx":0388
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6810
      Width           =   915
   End
   Begin FlexCell.Grid grdPOConfirm 
      Height          =   5145
      Left            =   60
      TabIndex        =   2
      Top             =   2520
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   9075
      BackColorBkg    =   -2147483633
      BackColorSel    =   -2147483633
      Cols            =   10
      DefaultFontSize =   8.25
      Rows            =   2
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   6990
      Picture         =   "frmPOConfirmation.frx":06C6
      Top             =   270
      Width           =   4695
   End
End
Attribute VB_Name = "frmPMISPOConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub InitGrid()
    With grdPOConfirm
        .Cols = 10: .Rows = 2
        .DisplayFocusRect = False: .AllowUserResizing = True

        '.BackColorFixed = &HFFCFB5            'RGB(90, 158, 214)
        '.BackColorFixedSel = &H8000000F       ' RGB(110, 180, 230) 'YELLOW
        '.BackColorBkg = &HF9EFE3              'RGB(90, 158, 214)
        '.BackColorScrollBar = RGB(231, 235, 247)
        '.BackColor1 = RGB(231, 235, 247)
        '.BackColor2 = RGB(239, 243, 255)
        '.GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = "Item No."
        .Cell(0, 2).Text = "Part No."
        .Cell(0, 3).Text = "Descrition"
        .Cell(0, 4).Text = "Ordered"
        .Cell(0, 5).Text = "Allocated"
        .Cell(0, 6).Text = "Fill"
        .Cell(0, 7).Text = "Kill"
        .Cell(0, 8).Text = "Status"
        .Cell(0, 9).Text = "Unit Price"

        .Column(1).CellType = cellTextBox
        .Column(2).CellType = cellTextBox:    '.Column(2).MaxLength = 50
        .Column(3).CellType = cellTextBox:    '.Column(3).MaxLength = 50
        .Column(4).CellType = cellTextBox
        .Column(5).CellType = cellTextBox
        .Column(6).CellType = cellCheckBox
        .Column(7).CellType = cellCheckBox
        .Column(8).CellType = cellTextBox
        .Column(9).CellType = cellTextBox

        .Column(0).Width = 1
        .Column(1).Width = 50: .Column(1).Locked = True
        .Column(2).Width = 100: .Column(2).Locked = True
        .Column(3).Width = 160: .Column(3).Locked = True
        .Column(4).Width = 50: .Column(4).Locked = True
        .Column(5).Width = 50: .Column(5).Locked = False
        .Column(6).Width = 20: .Column(6).Locked = False
        .Column(7).Width = 20: .Column(7).Locked = False
        .Column(8).Width = 50: .Column(8).Locked = True
        .Column(9).Width = 1: .Column(9).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 20
        '.Range(1, 16, .Rows - 1, 16).ForeColor = RGB(0, 0, 128)
    End With
End Sub

Sub FillDetails(XXX As String)
Dim rsPO_Details As ADODB.Recordset
Dim rsPO_TdayTran As ADODB.Recordset
Dim iJonathan As Integer
Dim vPOFill As Object
Dim vPOKill As Object
Set rsPO_Details = New ADODB.Recordset
Set rsPO_Details = gconDMIS.Execute("Select * from PMIS_vw_ConfirmedPO Where SONum = '" & txtSONum.Text & "'")
If Not rsPO_Details.EOF And Not rsPO_Details.BOF Then
    rsPO_Details.MoveFirst: iJonathan = 0
    Do While Not rsPO_Details.EOF
        iJonathan = iJonathan + 1
        grdPOConfirm.AddItem Null2String(rsPO_Details!itemno) & Chr(9) & _
                             Null2String(rsPO_Details!STOCK_ORD) & Chr(9) & _
                             SetPartDesc(Null2String(rsPO_Details!STOCK_ORD)) & Chr(9) & _
                             N2Str2Zero(rsPO_Details!Qty_Ordered) & Chr(9) & _
                             N2Str2Zero(rsPO_Details!Qty_Allocated) & Chr(9) & _
                             N2Str2Zero(rsPO_Details!POFill) & Chr(9) & _
                             N2Str2Zero(rsPO_Details!POKill) & Chr(9) & _
                             Null2String(rsPO_Details!Status) & Chr(9) & _
                             N2Str2Zero(rsPO_Details!UnitPrice)
        If iJonathan = 1 Then grdPOConfirm.RemoveItem 1
        rsPO_Details.MoveNext
    Loop
Else
    Set rsPO_TdayTran = New ADODB.Recordset
    Set rsPO_TdayTran = gconDMIS.Execute("Select * from PMIS_TDaytran Where Trantype = 'PO' and Tranno = '" & XXX & "' order by ItemNo asc")
    If Not rsPO_TdayTran.EOF And Not rsPO_TdayTran.BOF Then
        rsPO_TdayTran.MoveFirst: iJonathan = 0
        Do While Not rsPO_TdayTran.EOF
            iJonathan = iJonathan + 1
            grdPOConfirm.AddItem Null2String(rsPO_TdayTran!itemno) & Chr(9) & _
                                 Null2String(rsPO_TdayTran!STOCK_ORD) & Chr(9) & _
                                 SetPartDesc(Null2String(rsPO_TdayTran!STOCK_ORD)) & Chr(9) & _
                                 N2Str2Zero(rsPO_TdayTran!tranqty) & Chr(9) & _
                                 0 & Chr(9) & _
                                 0 & Chr(9) & _
                                 0 & Chr(9) & _
                                 "" & Chr(9) & _
                                 N2Str2Zero(rsPO_TdayTran!TRANINVAMT)
            If iJonathan = 1 Then grdPOConfirm.RemoveItem 1
            rsPO_TdayTran.MoveNext
        Loop
    End If
    Set rsPO_TdayTran = Nothing
End If
Set rsPO_Details = Nothing
End Sub

Private Sub cmdSavePOConfirmation_Click()
Dim kim As Integer

Dim vPO_NO As String
Dim vITEMNO As String
Dim vPartOrdered As String
Dim vDealerCode As String
Dim vSOMonth As String
Dim vSOYear As String
Dim vSONum As String
Dim vModelCode As String
Dim vSegment As String
Dim vQty_Ordered As Integer
Dim vQty_Allocated As Integer
Dim vQty_Unserved As Integer
Dim vFill As Integer
Dim vKill As Integer
Dim vQty_FillRate As Integer
Dim vQty_BackOrder As Integer
Dim vUnitPrice As Double
Dim vOrderAmount As Double
Dim vAllocAmount As Double
Dim vUnservedAmt As Double
Dim vAmtFillRate As Double
Dim vBackOrderAmt As Double
Dim vLineItemFillRate As Double
Dim vSOCategory As String
Dim vSOType As String
Dim vOrderScheme As String
Dim VStatus As String
Dim vPartsOrigin As String
Dim vByRegion As String

Dim rsPARTMAS As ADODB.Recordset
Dim rsPO_Details As ADODB.Recordset
picPOC.Visible = True
For kim = 1 To grdPOConfirm.Rows - 1
    
    Ordered = NumericVal(grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 4).Text)
    Served = NumericVal(grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 5).Text)
    vPartOrdered = N2Str2Null(grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 2).Text)
    vPO_NO = N2Str2Null(txtPONO.Text)
    vITEMNO = N2Str2Null(grdPOConfirm.Cell(kim, 1).Text)
    vDealerCode = N2Str2Null(txtDealerCode.Text)
    vSOMonth = N2Str2Null(Format(What_month(txtSOMonth.Text), "00"))
    vSOYear = N2Str2Null(txtSOYear.Text)
    vSONum = N2Str2Null(txtSONum.Text)
    vQty_Ordered = NumericVal(grdPOConfirm.Cell(kim, 4).Text)
    vQty_Allocated = NumericVal(grdPOConfirm.Cell(kim, 5).Text)
    vQty_Unserved = vQty_Ordered - vQty_Allocated
    vFill = NumericVal(grdPOConfirm.Cell(kim, 6).Text)
    vKill = NumericVal(grdPOConfirm.Cell(kim, 7).Text)
    vQty_FillRate = 0
    If vFill = 1 Then
        vQty_BackOrder = vQty_Unserved
    Else
        vQty_BackOrder = 0
    End If
    vUnitPrice = NumericVal(grdPOConfirm.Cell(kim, 9).Text)
    vOrderAmount = NumericVal(vQty_Ordered * vUnitPrice)
    vAllocAmount = NumericVal(vQty_Allocated * vUnitPrice)
    vUnservedAmt = NumericVal(vQty_Unserved * vUnitPrice)
    vAmtFillRate = 0
    vBackOrderAmt = NumericVal(vQty_BackOrder * vUnitPrice)
    vLineItemFillRate = 0
    vSOCategory = N2Str2Null(Left(txtSOType.Text, 1))
    VStatus = N2Str2Null(grdPOConfirm.Cell(kim, 8).Text)
    If NumericVal(grdPOConfirm.Cell(kim, 6).Text) = 1 Then
        vOrderScheme = "'F'"
    ElseIf NumericVal(grdPOConfirm.Cell(kim, 7).Text) = 1 Then
        vOrderScheme = "'K'"
    Else
        vOrderScheme = "NULL"
    End If
    If Left(txtSOType.Text, 1) = "W" Then vSOType = "'W'" Else vSOType = "'R'"
    Set rsPARTMAS = New ADODB.Recordset
    Set rsPARTMAS = gconDMIS.Execute("Select * from PMIS_PartMas Where PartNo = " & vPartOrdered)
    If Not rsPARTMAS.EOF And Not rsPARTMAS.BOF Then
        vModelCode = N2Str2Null(rsPARTMAS!modelcode)
        vSegment = N2Str2Null(rsPARTMAS!Segment)
        vPartsOrigin = N2Str2Null(rsPARTMAS!PartsOrigin)
        vByRegion = N2Str2Null(rsPARTMAS!Region)
    Else
        vModelCode = "NULL": vSegment = "NULL": vPartsOrigin = "NULL": vByRegion = "NULL"
    End If
    
    
    'PO_NO,ITEMNO,DealerCode,SOMonth,SOYear,SONum,Qty_Ordered,Qty_Allocated,Qty_Unserved
    'Fill,Kill,Qty_FillRate,Qty_BackOrder,UnitPrice,OrderAmount,AllocAmount
    'UnservedAmt,AmtFillRate,BackOrderAmt,LineItemFillRate,SOCategory
    'OrderScheme,vSOType,ModelCode,Segment,PartsOrigin,ByRegion
    
    Set rsPO_Details = New ADODB.Recordset
    Set rsPO_Details = gconDMIS.Execute("Select * from PMIS_PO_Details where SONum = " & vSONum & " and ItemNo = " & vITEMNO)
    If Not rsPO_Details.EOF And Not rsPO_Details.BOF Then
        gconDMIS.Execute ("Update PMIS_PO_Details Set" & _
                          " PO_NO = " & vPO_NO & ", ITEMNO = " & vITEMNO & "," & _
                          " DealerCode = " & vDealerCode & ", SOMonth = " & vSOMonth & "," & _
                          " SOYear = " & vSOYear & ", SONum = " & vSONum & "," & _
                          " Qty_Ordered = " & vQty_Ordered & ", Qty_Allocated = " & vQty_Allocated & "," & _
                          " Qty_Unserved = " & vQty_Unserved & ", POFill = " & vFill & "," & _
                          " POKill = " & vKill & ", Qty_FillRate = " & vQty_FillRate & "," & _
                          " Qty_BackOrder = " & vQty_BackOrder & ", UnitPrice = " & vUnitPrice & "," & _
                          " OrderAmount = " & vOrderAmount & ", AllocAmount = " & vAllocAmount & "," & _
                          " UnservedAmt = " & vUnservedAmt & ", AmtFillRate = " & vAmtFillRate & "," & _
                          " BackOrderAmt = " & vBackOrderAmt & ", LineItemFillRate = " & vLineItemFillRate & "," & _
                          " SOCategory = " & vSOCategory & ", OrderScheme = " & vOrderScheme & "," & _
                          " SOType = " & vSOType & ", ModelCode = " & vModelCode & "," & _
                          " Segment = " & vSegment & ", PartsOrigin = " & vPartsOrigin & "," & _
                          " ByRegion = " & vByRegion & ", Status = " & VStatus & _
                          " Where ID = " & rsPO_Details!ID)
    Else
        gconDMIS.Execute ("Insert Into PMIS_PO_Details " & _
                          "(PO_NO,ITEMNO,DealerCode,SOMonth,SOYear,SONum,Qty_Ordered,Qty_Allocated,Qty_Unserved," & _
                          "POFill,POKill,Qty_FillRate,Qty_BackOrder,UnitPrice,OrderAmount,AllocAmount," & _
                          "UnservedAmt,AmtFillRate,BackOrderAmt,LineItemFillRate,SOCategory,OrderScheme,SOType,ModelCode,Segment,Status,PartsOrigin,ByRegion) " & _
                          " values (" & vPO_NO & "," & vITEMNO & "," & vDealerCode & "," & vSOMonth & "," & vSOYear & "," & vSONum & "," & vQty_Ordered & "," & vQty_Allocated & "," & vQty_Unserved & _
                          "," & vFill & "," & vKill & "," & vQty_FillRate & "," & vQty_BackOrder & "," & vUnitPrice & "," & vOrderAmount & "," & vAllocAmount & _
                          "," & vUnservedAmt & "," & vAmtFillRate & "," & vBackOrderAmt & "," & vLineItemFillRate & "," & vSOCategory & "," & vOrderScheme & "," & vSOType & "," & vModelCode & "," & vSegment & "," & VStatus & "," & vPartsOrigin & "," & vByRegion & ")")
    End If
    progPOC.Value = (kim / (grdPOConfirm.Rows - 1)) * 100
    labPOC.Caption = progPOC.Value & " % Completed"
    DoEvents
Next
MsgBox "PO Confirmed...", vbInformation, "Info"
picPOC.Visible = False
Unload Me
End Sub

Private Sub cmdTranCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterMe Screen, Me, 0
picPOC.Visible = False: InitGrid
txtQty_Unserved = 0
txtQty_FillRate = 0
txtQty_BackOrder = 0
txtUnitPrice = ToDoubleNumber(0)
txtOrderAmount = ToDoubleNumber(0)
txtAllocAmount = ToDoubleNumber(0)
txtUnservedAmt = ToDoubleNumber(0)
txtBackOrderAmt = ToDoubleNumber(0)
End Sub

Sub ShowPODetails(XXX As String, YYY As String, ZZZ As String)
    Dim rsPO_Details As ADODB.Recordset
    Set rsPO_Details = New ADODB.Recordset
    Set rsPO_Details = gconDMIS.Execute("Select * from PMIS_PO_Details Where SONum = '" & XXX & "' AND ITEMNO = '" & YYY & "'")
    If Not rsPO_Details.EOF And Not rsPO_Details.BOF Then
        txtQty_Unserved = N2Str2Zero(rsPO_Details!Qty_Unserved)
        txtQty_FillRate = N2Str2Zero(rsPO_Details!Qty_FillRate)
        txtQty_BackOrder = N2Str2Zero(rsPO_Details!Qty_BackOrder)
        txtUnitPrice = ToDoubleNumber(N2Str2Zero(rsPO_Details!UnitPrice))
        txtOrderAmount = ToDoubleNumber(N2Str2Zero(rsPO_Details!OrderAmount))
        txtAllocAmount = ToDoubleNumber(N2Str2Zero(rsPO_Details!AllocAmount))
        txtUnservedAmt = ToDoubleNumber(N2Str2Zero(rsPO_Details!UnservedAmt))
        txtBackOrderAmt = ToDoubleNumber(N2Str2Zero(rsPO_Details!BackOrderAmt))
    Else
        Dim rsPO_TdayTran As ADODB.Recordset
        Set rsPO_TdayTran = New ADODB.Recordset
        Set rsPO_TdayTran = gconDMIS.Execute("Select * from PMIS_TDaytran Where TRANTYPE = 'PO' and TRANNO = '" & ZZZ & "' and ITEMNO = '" & YYY & "'")
        If Not rsPO_TdayTran.EOF And Not rsPO_TdayTran.BOF Then
            txtQty_Unserved = N2Str2Zero(rsPO_TdayTran!tranqty)
            txtQty_FillRate = 0
            txtQty_BackOrder = 0
            txtUnitPrice = ToDoubleNumber(N2Str2Zero(rsPO_TdayTran!TRANINVAMT))
            txtOrderAmount = ToDoubleNumber(N2Str2Zero(rsPO_TdayTran!tranqty) * N2Str2Zero(rsPO_TdayTran!TRANINVAMT))
            txtAllocAmount = ToDoubleNumber(0)
            txtUnservedAmt = ToDoubleNumber(N2Str2Zero(rsPO_TdayTran!tranqty) * N2Str2Zero(rsPO_TdayTran!TRANINVAMT))
            txtBackOrderAmt = ToDoubleNumber(0)
            SetPartsDetails Null2String(rsPO_TdayTran!STOCK_ORD)
        'Else
        '    txtQty_Unserved = 0
        '    txtQty_FillRate = 0
        '    txtQty_BackOrder = 0
        '    txtUnitPrice = ToDoubleNumber(0)
        '    txtOrderAmount = ToDoubleNumber(0)
        '    txtAllocAmount = ToDoubleNumber(0)
        '    txtUnservedAmt = ToDoubleNumber(0)
        '    txtBackOrderAmt = ToDoubleNumber(0)
        End If
        Set rsPO_TdayTran = Nothing
    End If
    Set rsPO_Details = Nothing
End Sub

Sub SetPartsDetails(XXX As String)
Dim rsPARTMAS As ADODB.Recordset
Set rsPARTMAS = New ADODB.Recordset
Set rsPARTMAS = gconDMIS.Execute("Select * from PMIS_Partmas Where Partno = '" & XXX & "'")
If Not rsPARTMAS.EOF And rsPARTMAS.BOF Then
    txtModelCode.Text = Null2String(rsPARTMAS!modelcode)
    txtSegment.Text = Null2String(rsPARTMAS!Segment)
    txtSOCategory.Text = Null2String(rsPARTMAS!SOCategory)
    txtPartsOrigin.Text = Null2String(rsPARTMAS!partorigin)
    txtByRegion.Text = Null2String(rsPARTMAS!ByRegion)
End If
Set rsPARTMAS = Nothing
End Sub

'Private Sub grdPOConfirm_CellChange(ByVal Row As Long, ByVal Col As Long)
'UpdatePOConfirmation
'End Sub

Private Sub grdPOConfirm_Click()
ShowPODetails Trim(txtSONum.Text), grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 1).Text, Trim(txtPONO.Text)
End Sub

Function SetPartDesc(ppp As String)
Set rsPMIS_Partmas = New ADODB.Recordset
    rsPMIS_Partmas.Open "Select partno,partdesc from PMIS_Partmas where partno = '" & ppp & "'", gconDMIS
If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
   SetPartDesc = UCase(Null2String(rsPMIS_Partmas!PARTDESC))
End If
End Function

Sub UpdatePOConfirmation()
ShowPODetails Trim(txtSONum.Text), grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 1).Text, Trim(txtPONO.Text)
Dim Ordered, Served As Double
Ordered = NumericVal(grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 4).Text)
Served = NumericVal(grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 5).Text)
If Served > Ordered Then
    MsgBox "Allocating Quantity more than Ordered Quantity is not Allowed..", vbInformation, "Not Allowed"
    grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 5).Text = 0
    Exit Sub
End If
If Served = Ordered Then
    'grdPOConfirm.Column(6).Locked = False
    'grdPOConfirm.Column(7).Locked = False
    'DoEvents
    'grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 6).BooleanValue = False
    'grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 7).BooleanValue = False
    'DoEvents
    grdPOConfirm.Column(6).Locked = True
    grdPOConfirm.Column(7).Locked = True
    txtQty_Unserved.Text = 0
    txtQty_FillRate.Text = 0
    txtQty_BackOrder.Text = 0
    txtAllocAmount = ToDoubleNumber(Served * NumericVal(txtUnitPrice.Text))
    txtUnservedAmt = ToDoubleNumber(0)
    txtBackOrderAmt = ToDoubleNumber(0)
    grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 8).Text = "Confirmed"
    Exit Sub
Else
    grdPOConfirm.Column(6).Locked = False
    grdPOConfirm.Column(7).Locked = False
    txtQty_Unserved.Text = Ordered - Served
    txtQty_FillRate.Text = 0
    If grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 6).IntegerValue = 1 Then
        txtQty_BackOrder.Text = Ordered - Served
    Else
        txtQty_BackOrder.Text = 0
    End If
    txtAllocAmount = ToDoubleNumber(Served * NumericVal(txtUnitPrice.Text))
    txtUnservedAmt = ToDoubleNumber(NumericVal(txtQty_Unserved.Text) * NumericVal(txtUnitPrice.Text))
    txtBackOrderAmt = ToDoubleNumber(NumericVal(txtQty_BackOrder.Text) * NumericVal(txtUnitPrice.Text))
End If
If grdPOConfirm.ActiveCell.Col = 6 Then
    grdPOConfirm.Column(6).Locked = False
    grdPOConfirm.Column(7).Locked = False
    'grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 6).BooleanValue = True
    'grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 7).BooleanValue = False
ElseIf grdPOConfirm.ActiveCell.Col = 7 Then
    grdPOConfirm.Column(6).Locked = False
    grdPOConfirm.Column(7).Locked = False
    'grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 7).BooleanValue = True
    'grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 6).BooleanValue = False
Else
End If
If grdPOConfirm.ActiveCell.Col = 6 Or grdPOConfirm.ActiveCell.Col = 7 Then
    If grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 6).BooleanValue = True Then
        grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 8).Text = "Confirmed"
    ElseIf grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 7).BooleanValue = True Then
        grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 8).Text = "Confirmed"
    Else
        grdPOConfirm.Cell(grdPOConfirm.ActiveCell.Row, 8).Text = ""
    End If
End If
End Sub

Private Sub grdPOConfirm_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
UpdatePOConfirmation
End Sub

Private Sub grdPOConfirm_KeyUp(KeyCode As Integer, Shift As Integer)
UpdatePOConfirmation
End Sub

Function SetPOUnitPrice(varPONO As String, varITEMNO As String)
Dim rsTDAYTRAN As ADODB.Recordset
Set rsTDAYTRAN = New ADODB.Recordset
Set rsTDAYTRAN = gconDMIS.Execute("Select * from PMIS_Tdaytran Where Trantype = 'PO' and Tranno = " & varPONO & " and ITEMNO = " & varITEMNO)
If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
    SetPOUnitPrice = N2Str2Zero(rsTDAYTRAN!TRANUPRICE)
End If
End Function
