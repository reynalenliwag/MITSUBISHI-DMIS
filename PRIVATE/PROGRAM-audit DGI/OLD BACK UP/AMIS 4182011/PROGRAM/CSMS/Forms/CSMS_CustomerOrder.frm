VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPMISCustomerOrder_CSMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Order"
   ClientHeight    =   6825
   ClientLeft      =   1110
   ClientTop       =   2520
   ClientWidth     =   11550
   ForeColor       =   &H00DEDFDE&
   Icon            =   "CSMS_CustomerOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   11550
   Begin VB.PictureBox fraAddTran 
      Height          =   3495
      Left            =   3720
      ScaleHeight     =   3435
      ScaleWidth      =   6825
      TabIndex        =   43
      Top             =   1290
      Width           =   6885
      Begin VB.Frame fraCostToCost 
         Height          =   405
         Left            =   2190
         TabIndex        =   122
         Top             =   1350
         Width           =   1575
         Begin VB.CheckBox Check1 
            Caption         =   "Cost to Cost"
            Height          =   195
            Left            =   90
            TabIndex        =   123
            Top             =   150
            Width           =   1395
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   3150
         TabIndex        =   121
         Top             =   1830
         Width           =   285
      End
      Begin VB.TextBox txtTranUCost 
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
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   116
         Text            =   "1000.00"
         ToolTipText     =   "Input price of item. Do not use comma and peso sign (e.g.300, 26)"
         Top             =   1440
         Width           =   945
      End
      Begin VB.Frame Frame2 
         Caption         =   "Parts Details"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3315
         Left            =   3870
         TabIndex        =   104
         Top             =   60
         Width           =   2865
         Begin VB.Frame Frame5 
            Caption         =   "Model Codes"
            Height          =   765
            Left            =   150
            TabIndex        =   119
            Top             =   2400
            Width           =   2595
            Begin VB.TextBox txtModelCode 
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
               Height          =   375
               Left            =   120
               MaxLength       =   6
               TabIndex        =   120
               ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
               Top             =   270
               Width           =   2325
            End
         End
         Begin VB.CheckBox chkAvailableOnStock 
            Alignment       =   1  'Right Justify
            Caption         =   "Available on Stock"
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
            Height          =   405
            Left            =   120
            TabIndex        =   118
            Top             =   270
            Width           =   2595
         End
         Begin VB.Frame Frame3 
            Height          =   975
            Left            =   150
            TabIndex        =   105
            Top             =   630
            Width           =   2595
            Begin VB.OptionButton optConsigned 
               Caption         =   "Consigned"
               Height          =   255
               Left            =   150
               TabIndex        =   108
               Top             =   660
               Width           =   1845
            End
            Begin VB.OptionButton optImported 
               Caption         =   "Imported"
               Height          =   255
               Left            =   150
               TabIndex        =   107
               Top             =   390
               Value           =   -1  'True
               Width           =   1845
            End
            Begin VB.OptionButton optLocalPurchase 
               Caption         =   "Local Purchases"
               Height          =   255
               Left            =   150
               TabIndex        =   106
               Top             =   150
               Width           =   1845
            End
         End
         Begin VB.Frame Frame4 
            Height          =   765
            Left            =   150
            TabIndex        =   109
            Top             =   1590
            Width           =   2595
            Begin VB.OptionButton optGenuine 
               Caption         =   "Genuine"
               Height          =   255
               Left            =   150
               TabIndex        =   111
               Top             =   180
               Value           =   -1  'True
               Width           =   1845
            End
            Begin VB.OptionButton optNonGenuine 
               Caption         =   "Non-Genuine"
               Height          =   255
               Left            =   150
               TabIndex        =   110
               Top             =   420
               Width           =   1845
            End
         End
      End
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
         Left            =   2880
         MouseIcon       =   "CSMS_CustomerOrder.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Delete Entry"
         Top             =   2580
         Width           =   735
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
         Left            =   2160
         MouseIcon       =   "CSMS_CustomerOrder.frx":0D47
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":0E99
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Cancel Entry"
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtTranDescription 
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
         Left            =   90
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1050
         Width           =   3675
      End
      Begin VB.TextBox txtTranTotalAmt 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   21
         Top             =   2160
         Width           =   1665
      End
      Begin VB.TextBox txtTranUPrice 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   20
         ToolTipText     =   "Input price of item. Do not use comma and peso sign (e.g.300, 26)"
         Top             =   1800
         Width           =   1665
      End
      Begin VB.TextBox txtTranQty 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   19
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   1440
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
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   16
         ToolTipText     =   "Type item number (e.g. 0001)"
         Top             =   60
         Width           =   615
      End
      Begin VB.ComboBox cboTranPartNo 
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
         ItemData        =   "CSMS_CustomerOrder.frx":11D7
         Left            =   1440
         List            =   "CSMS_CustomerOrder.frx":11D9
         Sorted          =   -1  'True
         TabIndex        =   17
         Text            =   "Combo1"
         ToolTipText     =   "Select Part Number from the list."
         Top             =   420
         Width           =   2295
      End
      Begin VB.TextBox txtPartID 
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
         Left            =   1470
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   420
         Width           =   585
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
         Left            =   1440
         MouseIcon       =   "CSMS_CustomerOrder.frx":11DB
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":132D
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Save Entry"
         Top             =   2580
         Width           =   735
      End
      Begin VB.Label labTranUCost 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost"
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
         Left            =   2250
         TabIndex        =   117
         Top             =   1470
         Width           =   615
      End
      Begin VB.Label labPartNo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Description"
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
         Left            =   1470
         TabIndex        =   58
         Top             =   1800
         Width           =   1275
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
         Left            =   1500
         TabIndex        =   56
         Top             =   1830
         Width           =   855
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Extend Price"
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
         TabIndex        =   50
         Top             =   2190
         Width           =   1305
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Left            =   840
         TabIndex        =   49
         Top             =   1830
         Width           =   615
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000D&
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
         Height          =   225
         Left            =   510
         TabIndex        =   48
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part No."
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
         Left            =   570
         TabIndex        =   47
         Top             =   450
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
         Left            =   570
         TabIndex        =   46
         Top             =   90
         Width           =   855
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000D&
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
         Height          =   225
         Left            =   120
         TabIndex        =   45
         Top             =   810
         Width           =   1275
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
         Left            =   1560
         TabIndex        =   57
         Top             =   1800
         Width           =   975
      End
   End
   Begin wizButton.cmd cmdAddTran 
      Height          =   3600
      Left            =   3660
      TabIndex        =   62
      Top             =   1230
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   6350
      TX              =   "cmd1"
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
      MICON           =   "CSMS_CustomerOrder.frx":167D
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   2640
      ScaleHeight     =   915
      ScaleWidth      =   8835
      TabIndex        =   92
      Top             =   5940
      Width           =   8835
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
         Left            =   7980
         MouseIcon       =   "CSMS_CustomerOrder.frx":1699
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":17EB
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   795
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
         Left            =   7200
         MouseIcon       =   "CSMS_CustomerOrder.frx":1B51
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":1CA3
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancelCO 
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
         Left            =   6420
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "CSMS_CustomerOrder.frx":2009
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":215B
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Cancel this Transaction"
         Top             =   0
         Width           =   795
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
         Left            =   5640
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "CSMS_CustomerOrder.frx":2495
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":25E7
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Post this Transaction"
         Top             =   0
         Width           =   795
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
         Left            =   4860
         MouseIcon       =   "CSMS_CustomerOrder.frx":290C
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":2A5E
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   795
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
         Left            =   4080
         MouseIcon       =   "CSMS_CustomerOrder.frx":2DBA
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":2F0C
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   795
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
         Left            =   3300
         MouseIcon       =   "CSMS_CustomerOrder.frx":321F
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":3371
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Move to Last Record"
         Top             =   0
         Width           =   795
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
         Left            =   2520
         MouseIcon       =   "CSMS_CustomerOrder.frx":36C1
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":3813
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Move to First Record"
         Top             =   0
         Width           =   795
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
         MouseIcon       =   "CSMS_CustomerOrder.frx":3B71
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":3CC3
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   795
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
         Left            =   960
         MouseIcon       =   "CSMS_CustomerOrder.frx":3FBD
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":410F
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   795
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
         Left            =   180
         MouseIcon       =   "CSMS_CustomerOrder.frx":4467
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":45B9
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9615
      ScaleHeight     =   885
      ScaleWidth      =   2220
      TabIndex        =   89
      Top             =   5895
      Width           =   2220
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
         Left            =   990
         MouseIcon       =   "CSMS_CustomerOrder.frx":4918
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":4A6A
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   795
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
         Left            =   210
         MouseIcon       =   "CSMS_CustomerOrder.frx":4DA8
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":4EFA
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   795
      End
   End
   Begin Crystal.CrystalReport rptCustomerOrder 
      Left            =   2820
      Top             =   4890
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Parts Issuance"
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
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2700
      ScaleHeight     =   255
      ScaleWidth      =   8685
      TabIndex        =   77
      Top             =   5520
      Width           =   8715
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Un-Post Transaction"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   6360
         TabIndex        =   82
         Top             =   30
         Width           =   2445
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 - Post Transaction"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   4380
         TabIndex        =   81
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Delete Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   2790
         TabIndex        =   80
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Edit Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1440
         TabIndex        =   79
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Add Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   90
         TabIndex        =   78
         Top             =   30
         Width           =   1455
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   6645
      Left            =   60
      TabIndex        =   69
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
         Left            =   75
         MaxLength       =   35
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   960
         Width           =   2445
      End
      Begin VB.OptionButton optRONo 
         Caption         =   "RO Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   71
         Top             =   630
         Width           =   2385
      End
      Begin VB.OptionButton optTranno 
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   70
         Top             =   390
         Value           =   -1  'True
         Width           =   2385
      End
      Begin MSComctlLib.ListView lstOrd_Hd 
         Height          =   5205
         Left            =   60
         TabIndex        =   73
         Top             =   1350
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   9181
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
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CSMS_CustomerOrder.frx":524A
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
      Begin VB.Label Label18 
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
         TabIndex        =   74
         Top             =   150
         Width           =   1455
      End
   End
   Begin SHDocVwCtl.WebBrowser browRIV 
      Height          =   2625
      Left            =   2820
      TabIndex        =   27
      Top             =   -2790
      Width           =   8565
      ExtentX         =   15108
      ExtentY         =   4630
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox picDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   2700
      ScaleHeight     =   2190
      ScaleWidth      =   8715
      TabIndex        =   42
      Top             =   3285
      Width           =   8745
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   8100
         Top             =   120
      End
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   2085
         Left            =   60
         TabIndex        =   15
         Top             =   60
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   3678
         _Version        =   393216
         Cols            =   7
         BackColor       =   16777215
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
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3165
      Left            =   2730
      ScaleHeight     =   3135
      ScaleWidth      =   8685
      TabIndex        =   28
      Top             =   90
      Width           =   8715
      Begin VB.Frame fraPayType 
         Caption         =   "Payment Type"
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
         Height          =   645
         Left            =   4560
         TabIndex        =   113
         Top             =   2430
         Width           =   4005
         Begin VB.OptionButton optCHARGE 
            Caption         =   "CHARGE"
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
            Left            =   2550
            TabIndex        =   115
            Top             =   240
            Width           =   1425
         End
         Begin VB.OptionButton optCASH 
            Caption         =   "CASH"
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
            Left            =   1530
            TabIndex        =   114
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.ComboBox cboRefPRSNo 
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2430
         TabIndex        =   8
         Text            =   "cboRefPRSNo"
         ToolTipText     =   "Select name of salesman from the list."
         Top             =   2370
         Width           =   1995
      End
      Begin VB.CommandButton Command1 
         Caption         =   "F2 - Select Customer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   75
         Top             =   960
         Width           =   2685
      End
      Begin VB.CommandButton c 
         Caption         =   "F1 - Assign PIS Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   84
         Top             =   60
         Width           =   2175
      End
      Begin VB.CommandButton cmdPISNum 
         Caption         =   "..."
         Height          =   375
         Left            =   7080
         TabIndex        =   83
         Top             =   60
         Width           =   255
      End
      Begin VB.TextBox txtReferencePIS 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   360
         Left            =   5160
         TabIndex        =   1
         Text            =   "PIWGC06H360"
         ToolTipText     =   "Type Reference PIS Number"
         Top             =   60
         Width           =   1935
      End
      Begin VB.ComboBox cboChargeTo 
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
         Height          =   330
         Left            =   5550
         TabIndex        =   11
         Text            =   "cboChargeTo"
         ToolTipText     =   "Select option from list."
         Top             =   -405
         Width           =   1785
      End
      Begin VB.TextBox txtRemarks 
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
         Height          =   615
         Left            =   4560
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         ToolTipText     =   "Type your message or remarks."
         Top             =   1740
         Width           =   4035
      End
      Begin VB.TextBox txtCustName 
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
         Height          =   945
         Left            =   60
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         ToolTipText     =   "Type complete name of customer."
         Top             =   1380
         Width           =   4365
      End
      Begin VB.TextBox txtTranDate 
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
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   3
         ToolTipText     =   "Type the date of transaction in mm/dd/yyyy format (e.g 7/5/2004)"
         Top             =   570
         Width           =   1545
      End
      Begin VB.TextBox txtDS1 
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
         Left            =   4800
         MaxLength       =   3
         TabIndex        =   12
         ToolTipText     =   "Type percentage to be added in the total amount. Do not include percent sign (e.g. 10, 15)"
         Top             =   945
         Width           =   525
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1740
         Picture         =   "CSMS_CustomerOrder.frx":53AC
         ScaleHeight     =   405
         ScaleWidth      =   435
         TabIndex        =   60
         Top             =   -540
         Width           =   435
         Begin VB.TextBox txtTranType 
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
            Left            =   0
            MaxLength       =   3
            TabIndex        =   61
            Top             =   60
            Width           =   525
         End
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
         Left            =   5700
         MaxLength       =   10
         TabIndex        =   13
         ToolTipText     =   "Input the type of the added amount."
         Top             =   945
         Width           =   1365
      End
      Begin VB.TextBox txtCustCode 
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
         Left            =   3420
         MaxLength       =   6
         TabIndex        =   6
         ToolTipText     =   "Input customer code (e.g. S01163)"
         Top             =   960
         Width           =   1005
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
         Left            =   3420
         MaxLength       =   7
         TabIndex        =   4
         ToolTipText     =   "Type the transaction terms."
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox txtRONO 
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
         Left            =   1170
         MaxLength       =   11
         TabIndex        =   5
         ToolTipText     =   "Type the transactin's RO number (e.g. A007541)"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtChargeTo 
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
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   -375
         Width           =   495
      End
      Begin VB.TextBox txtTranNo 
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
         Left            =   3420
         MaxLength       =   6
         TabIndex        =   0
         ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
         Top             =   60
         Width           =   1005
      End
      Begin VB.ComboBox cboSMName 
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
         Left            =   1080
         TabIndex        =   10
         ToolTipText     =   "Select name of salesman from the list."
         Top             =   2760
         Width           =   3345
      End
      Begin VB.ComboBox cboSalesMan 
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
         Height          =   330
         Left            =   1200
         TabIndex        =   9
         Text            =   "cboSalesMan"
         Top             =   1620
         Width           =   765
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1215
         Left            =   7110
         ScaleHeight     =   1215
         ScaleWidth      =   1515
         TabIndex        =   59
         Top             =   510
         Width           =   1515
         Begin VB.TextBox txtNetInvAmt 
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
            Left            =   90
            MaxLength       =   15
            TabIndex        =   66
            Top             =   810
            Width           =   1395
         End
         Begin VB.TextBox txtDS_Amt1 
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
            Left            =   90
            MaxLength       =   15
            TabIndex        =   65
            Top             =   440
            Width           =   1395
         End
         Begin VB.TextBox txtTTLInvAmt 
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
            Left            =   90
            MaxLength       =   15
            TabIndex        =   64
            Top             =   60
            Width           =   1395
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference PRS Number :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   90
         TabIndex        =   112
         Top             =   2400
         Width           =   2355
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PIS No."
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
         Left            =   4470
         TabIndex        =   76
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   4290
         TabIndex        =   68
         Top             =   120
         Width           =   165
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Height          =   285
         Left            =   5340
         TabIndex        =   67
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
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
         Left            =   5940
         TabIndex        =   31
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Man"
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
         TabIndex        =   41
         Top             =   2790
         Width           =   975
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
         Left            =   3840
         TabIndex        =   40
         Top             =   990
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL Amount"
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
         Left            =   5445
         TabIndex        =   39
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label Label7 
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
         Left            =   2250
         TabIndex        =   38
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label Label6 
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
         Left            =   2760
         TabIndex        =   37
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tran. Date"
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
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label labChargeTo 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Charge To"
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
         Left            =   4560
         TabIndex        =   35
         Top             =   -390
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tran. No."
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
         Left            =   2280
         TabIndex        =   34
         Top             =   90
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
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
         TabIndex        =   33
         Top             =   90
         Width           =   1725
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
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
         Left            =   4650
         TabIndex        =   32
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label labRONO 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "RO Number"
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
         TabIndex        =   30
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label labPosted 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "CANCELLED"
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
         Left            =   6900
         TabIndex        =   29
         Top             =   90
         Width           =   1725
      End
   End
   Begin wizButton.cmd cmdSignatories 
      Height          =   2535
      Left            =   4110
      TabIndex        =   63
      Top             =   1740
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   4471
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
      MICON           =   "CSMS_CustomerOrder.frx":80E8
   End
   Begin VB.PictureBox fraSignatories 
      Height          =   2355
      Left            =   4185
      ScaleHeight     =   2295
      ScaleWidth      =   4350
      TabIndex        =   51
      Top             =   1815
      Width           =   4410
      Begin VB.CommandButton cmdPrintRIV 
         Caption         =   "&Print PIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3030
         MouseIcon       =   "CSMS_CustomerOrder.frx":8104
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_CustomerOrder.frx":8256
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox chkPreview 
         BackColor       =   &H00DEDFDE&
         Height          =   255
         Left            =   4020
         TabIndex        =   26
         Top             =   1680
         Width           =   225
      End
      Begin VB.TextBox txtApprovedBy 
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
         Left            =   1440
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   780
         Width           =   2835
      End
      Begin VB.TextBox txtRequestedBy 
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
         Left            =   1440
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1140
         Width           =   2835
      End
      Begin VB.TextBox txtIssuedBy 
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
         Left            =   1440
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   420
         Width           =   2835
      End
      Begin VB.TextBox txtPreparedBy 
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
         Left            =   1440
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   60
         Width           =   2835
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Issued By"
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
         TabIndex        =   55
         Top             =   810
         Width           =   1395
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Received By"
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
         TabIndex        =   54
         Top             =   1140
         Width           =   1395
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
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
         TabIndex        =   53
         Top             =   420
         Width           =   1395
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By"
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
         TabIndex        =   52
         Top             =   90
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmPMISCustomerOrder_CSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOrd_Hd, rsTdayTran, rsPartMas                    As ADODB.Recordset
Attribute rsTdayTran.VB_VarUserMemId = 1073938432
Attribute rsPartMas.VB_VarUserMemId = 1073938432
Dim rsSalesMan, rsCunter, rsProfile                    As ADODB.Recordset
Attribute rsSalesMan.VB_VarUserMemId = 1073938435
Attribute rsCunter.VB_VarUserMemId = 1073938435
Attribute rsProfile.VB_VarUserMemId = 1073938435
Dim rsSignatories                                      As ADODB.Recordset
Attribute rsSignatories.VB_VarUserMemId = 1073938438
Dim rsREPOR, rsCUSTOMER                                As ADODB.Recordset
Attribute rsREPOR.VB_VarUserMemId = 1073938439
Attribute rsCUSTOMER.VB_VarUserMemId = 1073938439
Dim kcnt                                               As Integer
Attribute kcnt.VB_VarUserMemId = 1073938441
Dim AddorEdit                                          As String
Attribute AddorEdit.VB_VarUserMemId = 1073938442
Dim ORD_TOTUPRICE, ORD_TOTINVAMT, ORD_TOTVAT, ORD_TOTQTY As Double
Attribute ORD_TOTUPRICE.VB_VarUserMemId = 1073938443
Attribute ORD_TOTINVAMT.VB_VarUserMemId = 1073938443
Attribute ORD_TOTVAT.VB_VarUserMemId = 1073938443
Attribute ORD_TOTQTY.VB_VarUserMemId = 1073938443
Dim PrevOrdType, PrevOrdNo                             As String
Attribute PrevOrdType.VB_VarUserMemId = 1073938447
Attribute PrevOrdNo.VB_VarUserMemId = 1073938447
Dim LOCALACESS                                         As String


Private Sub cboChargeTo_LostFocus()
'If MsgBox("This transaction is CHARGE TO " & cboChargeTo.Text & vbCrLf & " Pls. Confirm...", vbYesNo + vbQuestion, "CHARGE TO " & cboChargeTo.Text) = vbNo Then
'cboChargeTo.SetFocus
'End If
End Sub

Private Sub cboRefPRSNo_Click()
cboRefPRSNo_LostFocus
End Sub

Private Sub cboRefPRSNo_GotFocus()
    Dim rsPRS                                          As ADODB.Recordset
    Dim rsPRS_Detail                                   As ADODB.Recordset
    Dim rsPRS_HDDup                             As ADODB.Recordset
    Set rsPRS = New ADODB.Recordset
    rsPRS.Open "Select tranno,refpisno from PMIS_vw_PRS WHERE [TYPE] = 'P' order by Tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPRS.EOF And Not rsPRS.BOF Then
        rsPRS.MoveFirst: cboRefPRSNo.Clear
        Do While Not rsPRS.EOF
            Set rsPRS_HDDup = New ADODB.Recordset
            rsPRS_HDDup.Open "select refpisno from PMIS_Ord_Hd where TRANTYPE <> 'PRS' AND [TYPE] = 'P' AND refprsno = '" & Null2String(rsPRS!refpisno) & "'", gconDMIS
        
            If Not rsPRS_HDDup.EOF And Not rsPRS_HDDup.BOF Then
            Else
               cboRefPRSNo.AddItem Null2String(rsPRS!refpisno)
            End If
            rsPRS.MoveNext
        Loop
    End If
End Sub

Private Sub cboRefPRSNo_LostFocus()
    If AddorEdit = "ADD" Then
        Dim rsRR_HDDup                             As ADODB.Recordset
        Set rsRR_HDDup = New ADODB.Recordset
        Dim rsCheckPartsOnHand As ADODB.Recordset
        rsRR_HDDup.Open "select refpisno,tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND refprsno = '" & cboRefPRSNo.Text & "'", gconDMIS
        If Not rsRR_HDDup.EOF And Not rsRR_HDDup.BOF Then
            MsgBox "PRS Number Already Received", vbInformation, "Invalid PRS Number"
            Exit Sub
        Else
            Set rsRR_HDDup = New ADODB.Recordset
            rsRR_HDDup.Open "select tranno,DS1 from PMIS_vw_PRS where [TYPE] = 'P' AND refpisno = '" & cboRefPRSNo.Text & "'", gconDMIS
            If Not rsRR_HDDup.EOF And Not rsRR_HDDup.BOF Then
                kcnt = 0: ORD_TOTUPRICE = 0: ORD_TOTINVAMT = 0: ORD_TOTVAT = 0: ORD_TOTQTY = 0
                Dim STOCKDESCription                               As String
                Set rsTdayTran = New ADODB.Recordset: cleargrid grdDetails
                'rsTdaytran.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " and trantype = " & N2Str2Null(rsOrd_Hd!trantype) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                rsTdayTran.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where [TYPE] = 'P' AND tranno = " & N2Str2Null(rsRR_HDDup!Tranno) & " and trantype = 'PRS' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
                    cboChargeTo.Enabled = False: Screen.MousePointer = 11: rsTdayTran.MoveFirst
                    Do While Not rsTdayTran.EOF
                        kcnt = kcnt + 1
                        'Set rsCheckPartsOnHand = New ADODB.Recordset
                        'Set rsCheckPartsOnHand = gconDMIS.Execute("Select OnHand from PMIS_StockMas where OnHand > 0 AND TYPE = 'P' and STOCKNO = " & N2Str2Null(rsTDAYTRAN!STOCK_ORD))
                        'If Not rsCheckPartsOnHand.EOF And Not rsCheckPartsOnHand.BOF Then
                            STOCKDESCription = SetSTOCKDESC(Null2String(rsTdayTran!STOCK_SUP))
                            grdDetails.AddItem rsTdayTran!ID & Chr(9) & Null2String(rsTdayTran!itemno) & Chr(9) & _
                                               Null2String(rsTdayTran!STOCK_ORD) & Chr(9) & _
                                               STOCKDESCription & Chr(9) & _
                                               N2Str2IntZero(rsTdayTran!tranqty) & Chr(9) & _
                                               Format(N2Str2Zero(rsTdayTran!TRANUPRICE), MAXIMUM_DIGIT) & Chr(9) & _
                                               Format(N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE), MAXIMUM_DIGIT)
                            ORD_TOTQTY = ORD_TOTQTY + N2Str2IntZero(rsTdayTran!tranqty)
                            ORD_TOTUPRICE = ORD_TOTUPRICE + (N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE))
                            ORD_TOTINVAMT = ORD_TOTINVAMT + (N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE))
                            rsTdayTran.MoveNext
                        'End If
                    Loop
                    txtTTLInvAmt.Text = ToDoubleNumber(ORD_TOTUPRICE)
                    If N2Str2Zero(rsRR_HDDup!ds1) <> 0 Then
                        txtDS1.Text = N2Str2Zero(rsRR_HDDup!ds1)
                        If txtDS_Desc1.Text = "" Then
                            txtDS_Desc1.Text = "DISCOUNT"
                        End If
                        txtDS_Amt1.Text = ToDoubleNumber(NumericVal(ORD_TOTUPRICE) * (NumericVal(txtDS1.Text) / 100))
                        txtNetInvAmt.Text = ToDoubleNumber(NumericVal(ORD_TOTUPRICE) - NumericVal(txtDS_Amt1.Text))
                    Else
                        txtDS1.Text = N2Str2Zero(rsRR_HDDup!ds1)
                        txtDS_Desc1.Text = ""
                        txtDS_Amt1.Text = "0.00"
                        txtNetInvAmt.Text = ToDoubleNumber(ORD_TOTUPRICE)
                    End If
                    ORD_TOTINVAMT = ORD_TOTINVAMT - NumericVal(txtDS_Amt1.Text)
                    If kcnt <> 0 Then grdDetails.RemoveItem 1
                    Screen.MousePointer = 0
                End If
            Else
                MsgSpeechBox "Invalid Parts Requisition Number!": If AddorEdit = "ADD" Then cleargrid grdDetails
            End If
        End If
    End If
End Sub

Private Sub cboSMName_Click()
    Set rsSalesMan = New ADODB.Recordset
    rsSalesMan.Open "select empno,signname from PMIS_vw_SalesMan where signname = " & N2Str2Null(cboSMName.Text), gconDMIS
    If Not rsSalesMan.EOF And Not rsSalesMan.BOF Then
        cboSalesMan.Text = Null2String(rsSalesMan!empno)
    End If
End Sub

Private Sub cboSMName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdSave.Value = True
End Sub

Private Sub cboTranPartNo_Change()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
        Check1.Enabled = True
    Else
        Check1.Enabled = False
    End If
End Sub

Private Sub cboTranPartNo_Click()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
        Check1.Enabled = True
    Else
        Check1.Enabled = False
    End If
End Sub

Private Sub cboTranPartNo_LostFocus()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
    End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   txtTranUPrice.Text = txtTranUCost.Text
Else
   txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
End If
End Sub

Private Sub cmdAddTran_Click()
    SendToBack
    cmdAddTran.Visible = True
    cmdAddTran.ZOrder 0
    fraAddTran.Visible = True
    fraAddTran.ZOrder 0
    fraAddTran.Enabled = True
    AddorEdit = "ADD"
    cmdTranDelete.Enabled = False
    InitParts
    On Error Resume Next
    cboTranPartNo.SetFocus
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", LOCALACESS) = False Then Exit Sub

    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    If LOGLEVEL <> "ADM" Then
        MsgBox "Warning: Your account is not allowed to cancel this transaction!", vbCritical, "Error"
        Exit Sub
    End If
    If MsgQuestionBox("Are you sure you want to Cancel this Transaction?", "Cancel Transaction") = True Then
        Dim PCurOnHand, PCurTISSQTY, PCurIssuances     As Integer
        Dim rsTdaytranDup, rsPartmasDup                As ADODB.Recordset

        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_TdayTran where [TYPE] = 'P' AND tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " and trantype = " & N2Str2Null(rsOrd_Hd!TRANTYPE), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
            rsTdaytranDup.MoveFirst
            Do While Not rsTdaytranDup.EOF
                Set rsPartmasDup = New ADODB.Recordset
                rsPartmasDup.Open "select STOCKNO,onhand,tissqty,TISSQTY,issuances,REQSERVED,S_REQSERVED from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD) & " AND ACTIVE = 'Y'", gconDMIS
                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                    PCurOnHand = N2Str2IntZero(rsPartmasDup!ONHAND) + N2Str2Zero(rsTdaytranDup!tranqty)
                    PCurTISSQTY = N2Str2IntZero(rsPartmasDup!TISSQTY) - N2Str2Zero(rsTdaytranDup!tranqty)
                    PCurIssuances = N2Str2IntZero(rsPartmasDup!issuances) - N2Str2Zero(rsTdaytranDup!tranqty)
                    If Null2String(rsOrd_Hd!Status) = "P" Then
                        If Null2String(rsOrd_Hd!SALES_ORIGIN) = "W" Then
                            gconDMIS.Execute "update PMIS_PARTMAS set" & _
                                           " REQSERVED = " & N2Str2IntZero(rsPartmasDup!REQServed) - N2Str2Zero(rsTdaytranDup!tranqty) & _
                                           " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                        Else
                            gconDMIS.Execute "update PMIS_PARTMAS set" & _
                                           " S_REQSERVED = " & N2Str2IntZero(rsPartmasDup!S_REQServed) - N2Str2Zero(rsTdaytranDup!tranqty) & _
                                           " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                        End If
                        gconDMIS.Execute "update PMIS_PARTMAS set" & _
                                       " onhand = " & PCurOnHand & "," & _
                                       " tissqty = " & PCurTISSQTY & "," & _
                                       " issuances = " & PCurIssuances & "," & _
                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                       " lastupdate = '" & LOGDATE & "'" & _
                                       " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                    End If
                    gconDMIS.Execute "update PMIS_TdayTran set" & _
                                   " status = 'C'," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where id = " & rsTdaytranDup!ID
                End If
                rsTdaytranDup.MoveNext
            Loop
        End If
        gconDMIS.Execute "update PMIS_Ord_Hd set" & _
                       " status = 'C'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labID.Caption
        rsRefresh
        LogAudit "C", "CUSTOMER ORDER", txtTranNo & txtCustCode
        On Error Resume Next
        rsOrd_Hd.Find "id =" & labID.Caption
        StoreMemVars
    End If
    Set rsTdaytranDup = Nothing
    Set rsPartmasDup = Nothing

    Exit Sub
Errorcode:
    ShowVBError

End Sub

'Private Sub cmdPISNum_Click()
'    With frmPMISPIFormation
'        If AddorEdit = "EDIT" Then
'            .txtedit = "EDIT"
'        Else
'            .txtedit = ""
'        End If
'        .lbl2 = Mid(txtReferencePIS, 3, 1)
'        .lbl3 = Mid(txtReferencePIS, 4, 1)
'        .lbl4 = Mid(txtReferencePIS, 5, 1)
'        ' .lbl6_7 = Mid(txtReferencePIS, 6, 2)
'        ' .lbl8 = Mid(txtReferencePIS, 8, 1)
'        .lbl9.Text = Mid(txtReferencePIS, 9, 3)
'        .lbl11 = Mid(txtReferencePIS, 12, 1)
'        If .lbl2.Caption = "S" Then
'            .optS.Value = True
'        ElseIf .lbl2.Caption = "W" Then
'            .optW.Value = True
'        ElseIf .lbl2.Caption = "M" Then
'            .optM.Value = True
'        ElseIf .lbl2.Caption = "J" Then
'            .optJ.Value = True
'        ElseIf .lbl2.Caption = "O" Then
'            .optO.Value = True
'        End If
'        If .lbl3.Caption = "G" Then
'            .optG.Value = True
'        ElseIf .lbl3.Caption = "B" Then
'            .optB.Value = True
'        End If
'        If .lbl4.Caption = "C" Then
'            .optC.Value = True
'        ElseIf .lbl4.Caption = "I" Then
'            .optI.Value = True
'        ElseIf .lbl4.Caption = "W" Then
'            .optW2.Value = True
'        End If
'        If .lbl11.Caption = "1" Then
'            .opt1.Value = True
'        ElseIf .lbl11.Caption = "2" Then
'            .opt2.Value = True
'        ElseIf .lbl11.Caption = "0" Then
'            .opt0.Value = True
'        End If
'    End With
'    frmPMISPIFormation.Show 1
'    On Error Resume Next
'    txtCustName.SetFocus
'End Sub

'Private Sub cmdPost_Click()
'    If Function_Access(LOGID, "Acess_Post", LOCALACESS) = False Then Exit Sub
'
'    'updating code:    JAA - 07112007
'    On Error GoTo Errorcode:
'
'    'MsgSpeechBox "Posting of Transaction is Automated by OR System or Billing System." & vbCrLf & _
'     '             "Manual Posting is made only by your System Administrator."
'    If MsgQuestionBox("Are you sure you want to Post this Transaction?", "Post Transaction") = True Then
'        Dim PCurOnHand, PCurTISSQTY, PCurIssuances     As Integer
'        Dim rsTdaytranDup, rsPartmasDup                As ADODB.Recordset
'
'        Set rsTdaytranDup = New ADODB.Recordset
'        rsTdaytranDup.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_TdayTran where [TYPE] = 'P' AND tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " and trantype = " & N2Str2Null(rsOrd_Hd!TRANTYPE), gconDMIS, adOpenForwardOnly, adLockReadOnly
'        If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
'            rsTdaytranDup.MoveFirst
'            Do While Not rsTdaytranDup.EOF
'                Set rsPartmasDup = New ADODB.Recordset
'                rsPartmasDup.Open "select STOCKNO,onhand,tissqty,TISSQTY,issuances,REQSERVED,S_REQSERVED from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD) & " AND ACTIVE = 'Y'", gconDMIS
'                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
'                    PCurOnHand = N2Str2IntZero(rsPartmasDup!ONHAND) - N2Str2Zero(rsTdaytranDup!tranqty)
'                    PCurTISSQTY = N2Str2IntZero(rsPartmasDup!TISSQTY) + N2Str2Zero(rsTdaytranDup!tranqty)
'                    PCurIssuances = N2Str2IntZero(rsPartmasDup!issuances) + N2Str2Zero(rsTdaytranDup!tranqty)
'                    If Null2String(rsOrd_Hd!SALES_ORIGIN) = "W" Then
'                        gconDMIS.Execute "update PMIS_PARTMAS set" & _
'                                       " REQSERVED = " & N2Str2IntZero(rsPartmasDup!REQServed) + N2Str2Zero(rsTdaytranDup!tranqty) & _
'                                       " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
'                    Else
'                        gconDMIS.Execute "update PMIS_PARTMAS set" & _
'                                       " S_REQSERVED = " & N2Str2IntZero(rsPartmasDup!S_REQServed) + N2Str2Zero(rsTdaytranDup!tranqty) & _
'                                       " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
'                    End If
'                    gconDMIS.Execute "update PMIS_PARTMAS set" & _
'                                   " onhand = " & PCurOnHand & "," & _
'                                   " tissqty = " & PCurTISSQTY & "," & _
'                                   " issuances = " & PCurIssuances & "," & _
'                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
'                                   " lastupdate = '" & LOGDATE & "'" & _
'                                   " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
'                    gconDMIS.Execute "update PMIS_TdayTran set" & _
'                                   " status = 'P'," & _
'                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
'                                   " lastupdate = '" & LOGDATE & "'" & _
'                                   " where id = " & rsTdaytranDup!ID
'                End If
'                rsTdaytranDup.MoveNext
'            Loop
'        End If
'        gconDMIS.Execute "update PMIS_Ord_Hd set" & _
'                       " status = 'P'," & _
'                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
'                       " lastupdate = '" & LOGDATE & "'" & _
'                       " where id = " & labID.Caption
'        LogAudit "P", "CUSTOMER ORDER", txtTranNo & txtCustCode
'        rsRefresh
'        On Error Resume Next
'
'        rsOrd_Hd.Find "id =" & labID.Caption
'        StoreMemVars
'    End If
'    Set rsTdaytranDup = Nothing
'    Set rsPartmasDup = Nothing
'    'If txtTranType.Text = "RIV" Then ImportParts txtRONO
'    Exit Sub
'Errorcode:
'    ShowVBError
'
'End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", LOCALACESS) = False Then Exit Sub

    'updating code:    JAA - 07112007
    'On Error GoTo Errorcode:

    'If rsOrd_Hd!TRANTYPE = "RIV" Then
    '    If MsgQuestionBox("Parts Issuance Slip will be Printed. Are you Sure?", "Confirm Printing...") = True Then RIVPRINTING
    'End If
    If rsOrd_Hd!TRANTYPE = "ADB" Or rsOrd_Hd!TRANTYPE = "RIV" Then
        cmdSignatories.Visible = True
        cmdSignatories.ZOrder 0
        fraSignatories.Visible = True
        fraSignatories.ZOrder 0
'        Dim rsSignatories As ADODB.Recordset
'        Set rsSignatories = New ADODB.Recordset
'        rsSignatories.Open "Select * from ALL_Profile where modulename = 'PMIS'", gconDMIS
'        If Not rsSignatories.EOF And Not rsSignatories.BOF Then
'            '=========================================================
'            'updating code:     JAA - 09252007
'            'txtPreparedBy.Text = Null2String(rsSignatories!preparedby)
'            'txtIssuedBy.Text = Null2String(rsSignatories!issuedby)
'            'txtApprovedBy.Text = Null2String(rsSignatories!approvedby)
'            '=========================================================
'            On Error Resume Next
'            txtRequestedBy.SetFocus
'        End If
'        Set rsSignatories = Nothing
            txtPreparedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", "")
            txtIssuedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", "")
            txtApprovedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", "")
            On Error Resume Next
            txtRequestedBy.SetFocus
    End If
    If rsOrd_Hd!TRANTYPE = "CSH" Then
        If MsgQuestionBox("Parts Issuance Slip (CSH) will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            CSHPRINTING
        End If
    End If
    If rsOrd_Hd!TRANTYPE = "CHG" Then
        If MsgQuestionBox("Parts Issuance Slip (CHG) will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            CHGPRINTING
        End If
    End If
    If rsOrd_Hd!TRANTYPE = "DR" Then
        If MsgQuestionBox("DR Out Transaction will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            'DRPRINTING
            'UPDATED BY: FML (07/23/2007) - CUSTOMIZED DR PRINTING FOR HAI
            'NEWDRPRINTING
        End If
    End If

    LogAudit "V", "PARTS CUSTOMER ORDER", txtTranNo & txtCustCode

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Sub CHGPRINTING()
    Dim Filter                                         As String
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHG.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHGDisc.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
End Sub

Sub CSHPRINTING()
    Dim Filter                                         As String
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH.RPT", "{ord_hd.TYPE} = 'P' AND {ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSHDisc.RPT", "{ord_hd.TYPE} = 'P' AND {ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
End Sub

Sub PISPRINTING()
    Dim Filter                                         As String
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "PIS.RPT", "{ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "PISDisc.RPT", "{ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
End Sub

'Sub SERVICEPISPRINTING()
'    Screen.MousePointer = 11
'    Dim cnt1, cnt2, cnt3                               As Integer
'    Dim knt, cntCOPY                                   As Integer
'    Dim TOTALQTY, TOTALPRICE                           As Double
'    Dim Filter                                         As String
'    Set rsProfile = New ADODB.Recordset
'    rsProfile.Open "select * from ALL_Profile where ModuleName = 'PMIS'", gconDMIS
'    Open App.Path & "\PIS.HTML" For Output As #1
'    Set rsTdayTran = New ADODB.Recordset
'    rsTdayTran.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where TYPE = 'P' AND tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " and trantype = 'RIV' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
'    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
'        TOTALQTY = 0
'        TOTALPRICE = 0
'        If rsTdayTran.RecordCount > MAX_ISS_LINE Then cntCOPY = 4 Else cntCOPY = 2
'        Print #1, "<html><body>"
'        knt = 0
'        For knt = 1 To cntCOPY
'            If knt < 3 Then
'                rsTdayTran.MoveFirst
'                TOTALQTY = 0: TOTALPRICE = 0
'            Else
'                If rsTdayTran.EOF Then
'                    rsTdayTran.MoveLast
'                Else
'                    rsTdayTran.MoveNext
'                End If
'            End If
'            Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
'            Print #1, "<tr>"
'            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
'            Print #1, "<td align=center width=60%><font size=3 FACE=TIMES NEW ROMAN>" & rsProfile!CompanyName & "</font></td>"
'            Print #1, "<td align=right width=20%><font size=1 FACE=TIMES NEW ROMAN>COPY: " & knt & "</font></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
'            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>PARTS ISSUANCE SLIP</strong></font></td>"
'            Print #1, "<td align=left width=20%>&nbsp;</td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td align=left width=20%>&nbsp;</td>"
'            Print #1, "<td align=center width=60%>&nbsp;</td>"
'            Print #1, "<td align=left width=20%>&nbsp;</td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & "SERVICE PIS-" & Null2String(rsOrd_Hd!Tranno) & "</b></i></u></FONT></td>"
'            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(rsOrd_Hd!trandate) & "</b></FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(rsOrd_Hd!custcode) & "</b></FONT></td>"
'            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(rsOrd_Hd!custname) & "</b></FONT></td>"
'            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=5%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ITM #</FONT></td>"
'            Print #1, "<td width=20%><FONT SIZE=2 FACE=TIMES NEW ROMAN>PART NUMBER</FONT></td>"
'            Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>DESCRIPTION</FONT></td>"
'            Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>QTY</FONT></td>"
'            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>UNIT PRICE</FONT></td>"
'            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>TOTAL PRICE</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
'            cnt1 = 0
'            If rsTdayTran.RecordCount > MAX_ISS_LINE Then
'                cnt2 = 0
'            Else
'                cnt2 = MAX_ISS_LINE - rsTdayTran.RecordCount
'            End If
'            If knt >= 3 Then cnt2 = MAX_ISS_LINE - (rsTdayTran.RecordCount - MAX_ISS_LINE)
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            If rsTdayTran.AbsolutePosition > MAX_ISS_LINE Then
'                rsTdayTran.AbsolutePosition = MAX_ISS_LINE + 1
'            End If
'            Do While Not rsTdayTran.EOF
'                Print #1, "<tr>"
'                Print #1, "<td width=5%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(rsTdayTran!itemno) & "</FONT></td>"
'                Print #1, "<td width=20%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(rsTdayTran!STOCK_ORD) & "</FONT></td>"
'                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & SetSTOCKDESC(Null2String(rsTdayTran!STOCK_SUP)) & "</FONT></td>"
'                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(rsTdayTran!tranqty) & "</FONT></td>"
'                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(rsTdayTran!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
'                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
'                If knt <> 4 Then
'                    TOTALQTY = TOTALQTY + N2Str2IntZero(rsTdayTran!tranqty)
'                    TOTALPRICE = TOTALPRICE + N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE)
'                End If
'                Print #1, "</tr>"
'                If rsTdayTran.AbsolutePosition = MAX_ISS_LINE Then Exit Do
'                rsTdayTran.MoveNext
'            Loop
'            For cnt3 = 1 To cnt2
'                Print #1, "<tr>"
'                Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "</tr>"
'            Next
'            Print #1, "</table>"
'            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
'            If cntCOPY = 4 And knt < 3 Then
'                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'                Print #1, "<tr>"
'                Print #1, "<td width=5%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=20%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=35%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=8%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "</tr>"
'                Print #1, "</table>"
'            Else
'                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'                Print #1, "<tr>"
'                Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL PIS</FONT></td>"
'                Print #1, "<td align=right width=8%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & TOTALQTY & "</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td align=right width=17%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & Format(TOTALPRICE, MAXIMUM_DIGIT) & "</FONT></td>"
'                Print #1, "</tr>"
'                Print #1, "</table>"
'            End If
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtPreparedBy.Text & "</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtIssuedBy.Text & "</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtApprovedBy.Text & "</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtRequestedBy.Text & "</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Requested By</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Approved By</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Issued By</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Received By</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            If knt <> 2 And knt <> 4 Then
'                Print #1, "<table>"
'                Print #1, "<tr>"
'                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'                Print #1, "</tr>"
'                Print #1, "</table>"
'                Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
'                Print #1, "<table>"
'                Print #1, "<tr>"
'                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'                Print #1, "</tr>"
'                Print #1, "</table>"
'            End If
'        Next
'        Print #1, "</body></html>"
'        Close #1
'        On Error Resume Next
'        Open App.Path & "\PIS.HTML" For Input As #1
'        If EOF(1) Then
'            MsgSpeechBox "File Not Found!"
'            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
'        Else
'            Close #1
'            browRIV.Navigate "about:blank"
'            browRIV.Refresh
'            browRIV.Navigate App.Path & "\PIS.HTML"
'            DoEvents
'            'If chkPreview.Value = 1 Then
'                browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
'            'Else
'            '    browRIV.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
'            'End If
'            Screen.MousePointer = 0
'        End If
'    End If
'    Set rsProfile = Nothing
'    Screen.MousePointer = 0
'End Sub

Sub SERVICEPISPRINTING()
    Screen.MousePointer = 11
    PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Parts.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'RIV' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub


Sub DRPRINTING()
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "DR.RPT", "{ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
End Sub

'Sub NEWDRPRINTING()
'    Screen.MousePointer = 11
'    Dim cnt1, cnt2, cnt3                               As Integer
'    Dim knt, cntCOPY                                   As Integer
'    Dim TOTALQTY, TOTALPRICE                           As Double
'    Dim Filter                                         As String
'    Set rsProfile = New ADODB.Recordset
'    rsProfile.Open "select * from ALL_Profile where ModuleName = 'PMIS'", gconDMIS
'    Open App.Path & "\DR.HTML" For Output As #1
'    Set rsTdayTran = New ADODB.Recordset
'    rsTdayTran.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where TYPE = 'P' and tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " and trantype = 'DR' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
'    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
'        TOTALQTY = 0
'        TOTALPRICE = 0
'        If rsTdayTran.RecordCount > MAX_ISS_LINE Then cntCOPY = 4 Else cntCOPY = 2
'        Print #1, "<html><body>"
'        knt = 0
'        For knt = 1 To cntCOPY
'            If knt < 3 Then
'                rsTdayTran.MoveFirst
'                TOTALQTY = 0: TOTALPRICE = 0
'            Else
'                If rsTdayTran.EOF Then
'                    rsTdayTran.MoveLast
'                Else
'                    rsTdayTran.MoveNext
'                End If
'            End If
'            Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
'            Print #1, "<tr>"
'            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
'            Print #1, "<td align=center width=60%><font size=3 FACE=TIMES NEW ROMAN>" & rsProfile!CompanyName & "</font></td>"
'            Print #1, "<td align=right width=20%><font size=1 FACE=TIMES NEW ROMAN>COPY: " & knt & "</font></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
'            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>PARTS DELIVERY RECEIPT</strong></font></td>"
'            Print #1, "<td align=left width=20%>&nbsp;</td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td align=left width=20%>&nbsp;</td>"
'            Print #1, "<td align=center width=60%>&nbsp;</td>"
'            Print #1, "<td align=left width=20%>&nbsp;</td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & Null2String(rsOrd_Hd!TRANTYPE) & "-" & Null2String(rsOrd_Hd!Tranno) & "</b></i></u></FONT></td>"
'            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(rsOrd_Hd!trandate) & "</b></FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(rsOrd_Hd!custcode) & "</b></FONT></td>"
'            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(rsOrd_Hd!custname) & "</b></FONT></td>"
'            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=10%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ITM #</FONT></td>"
'            Print #1, "<td width=15%><FONT SIZE=2 FACE=TIMES NEW ROMAN>PART NUMBER</FONT></td>"
'            Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>DESCRIPTION</FONT></td>"
'            Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>QTY</FONT></td>"
'            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>UNIT PRICE</FONT></td>"
'            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>TOTAL PRICE</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
'            cnt1 = 0
'            If rsTdayTran.RecordCount > MAX_ISS_LINE Then
'                cnt2 = 0
'            Else
'                cnt2 = MAX_ISS_LINE - rsTdayTran.RecordCount
'            End If
'            If knt >= 3 Then cnt2 = MAX_ISS_LINE - (rsTdayTran.RecordCount - MAX_ISS_LINE)
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            If rsTdayTran.AbsolutePosition > MAX_ISS_LINE Then
'                rsTdayTran.AbsolutePosition = MAX_ISS_LINE + 1
'            End If
'            Do While Not rsTdayTran.EOF
'                Print #1, "<tr>"
'                Print #1, "<td width=10%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(rsTdayTran!itemno) & "</FONT></td>"
'                Print #1, "<td width=15%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(rsTdayTran!STOCK_ORD) & "</FONT></td>"
'                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & SetSTOCKDESC(Null2String(rsTdayTran!STOCK_SUP)) & "</FONT></td>"
'                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(rsTdayTran!tranqty) & "</FONT></td>"
'                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(rsTdayTran!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
'                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
'                If knt <> 4 Then
'                    TOTALQTY = TOTALQTY + N2Str2IntZero(rsTdayTran!tranqty)
'                    TOTALPRICE = TOTALPRICE + N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE)
'                End If
'                Print #1, "</tr>"
'                If rsTdayTran.AbsolutePosition = MAX_ISS_LINE Then Exit Do
'                rsTdayTran.MoveNext
'            Loop
'            For cnt3 = 1 To cnt2
'                Print #1, "<tr>"
'                Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "</tr>"
'            Next
'            Print #1, "</table>"
'            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
'            If cntCOPY = 4 And knt < 3 Then
'                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'                Print #1, "<tr>"
'                Print #1, "<td width=10%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=15%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=35%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=8%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "</tr>"
'                Print #1, "</table>"
'            Else
'                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'                Print #1, "<tr>"
'                Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL DR</FONT></td>"
'                Print #1, "<td align=right width=8%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & TOTALQTY & "</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td align=right width=17%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & Format(TOTALPRICE, MAXIMUM_DIGIT) & "</FONT></td>"
'                Print #1, "</tr>"
'                Print #1, "</table>"
'            End If
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtPreparedBy.Text & "</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtIssuedBy.Text & "</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtApprovedBy.Text & "</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtRequestedBy.Text & "</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Requested By</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Approved By</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Issued By</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Received By</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            If knt <> 2 And knt <> 4 Then
'                Print #1, "<table>"
'                Print #1, "<tr>"
'                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'                Print #1, "</tr>"
'                Print #1, "</table>"
'                Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
'                Print #1, "<table>"
'                Print #1, "<tr>"
'                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'                Print #1, "</tr>"
'                Print #1, "</table>"
'            End If
'        Next
'        Print #1, "</body></html>"
'        Close #1
'        On Error Resume Next
'        Open App.Path & "\DR.HTML" For Input As #1
'        If EOF(1) Then
'            MsgSpeechBox "File Not Found!"
'            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
'        Else
'            Close #1
'            browRIV.Navigate "about:blank"
'            browRIV.Refresh
'            browRIV.Navigate App.Path & "\DR.HTML"
'            DoEvents
'            'If chkPreview.Value = 1 Then
'                browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
'            'Else
'            '    browRIV.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
'            'End If
'            Screen.MousePointer = 0
'        End If
'    End If
'    Set rsProfile = Nothing
'    Screen.MousePointer = 0
'End Sub

'Sub ADBPRINTING()
'    Screen.MousePointer = 11
'    Dim cnt1, cnt2, cnt3                               As Integer
'    Dim knt, cntCOPY                                   As Integer
'    Dim TOTALQTY, TOTALPRICE                           As Double
'    Dim Filter                                         As String
'    Set rsProfile = New ADODB.Recordset
'    rsProfile.Open "select * from ALL_Profile", gconDMIS
'    Open PMIS_REPORT_PATH & "ADB.HTML" For Output As #1
'    Set rsTdayTran = New ADODB.Recordset
'    rsTdayTran.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where [TYPE] = 'P' AND tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " and trantype = 'ADB' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
'    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
'        TOTALQTY = 0
'        TOTALPRICE = 0
'        If rsTdayTran.RecordCount > MAX_ISS_LINE Then
'            cntCOPY = 4
'        Else
'            cntCOPY = 2
'        End If
'        Print #1, "<html><body>"
'        knt = 0
'        For knt = 1 To cntCOPY
'            If knt < 3 Then
'                rsTdayTran.MoveFirst
'                TOTALQTY = 0: TOTALPRICE = 0
'            Else
'                If rsTdayTran.EOF Then
'                    rsTdayTran.MoveLast
'                Else
'                    rsTdayTran.MoveNext
'                End If
'            End If
'            Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
'            Print #1, "<tr>"
'            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
'            Print #1, "<td align=center width=60%><font size=3 FACE=TIMES NEW ROMAN>" & rsProfile!CompanyName & "</font></td>"
'            Print #1, "<td align=right width=20%><font size=1 FACE=TIMES NEW ROMAN>COPY: " & knt & "</font></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
'            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>ADVANCED BILL VOUCHER</strong></font></td>"
'            Print #1, "<td align=left width=20%>&nbsp;</td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td align=left width=20%>&nbsp;</td>"
'            Print #1, "<td align=center width=60%>&nbsp;</td>"
'            Print #1, "<td align=left width=20%>&nbsp;</td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & Null2String(rsOrd_Hd!TRANTYPE) & "-" & Null2String(rsOrd_Hd!Tranno) & "</b></i></u></FONT></td>"
'            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(rsOrd_Hd!trandate) & "</b></FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(rsOrd_Hd!custcode) & "</b></FONT></td>"
'            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Charge To: " & Null2String(rsOrd_Hd!chargeto) & "</b></FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(rsOrd_Hd!custname) & "</b></FONT></td>"
'            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Ref RO# : " & Null2String(rsOrd_Hd!rono) & "</b></FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=10%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ITM #</FONT></td>"
'            Print #1, "<td width=15%><FONT SIZE=2 FACE=TIMES NEW ROMAN>PART NUMBER</FONT></td>"
'            Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>DESCRIPTION</FONT></td>"
'            Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>QTY</FONT></td>"
'            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>UNIT PRICE</FONT></td>"
'            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>TOTAL PRICE</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
'            cnt1 = 0
'            If rsTdayTran.RecordCount > MAX_ISS_LINE Then
'                cnt2 = 0
'            Else
'                cnt2 = MAX_ISS_LINE - rsTdayTran.RecordCount
'            End If
'            If knt >= 3 Then cnt2 = MAX_ISS_LINE - (rsTdayTran.RecordCount - MAX_ISS_LINE)
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            If rsTdayTran.AbsolutePosition > MAX_ISS_LINE Then
'                rsTdayTran.AbsolutePosition = MAX_ISS_LINE + 1
'            End If
'            Do While Not rsTdayTran.EOF
'                Print #1, "<tr>"
'                Print #1, "<td width=10%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(rsTdayTran!itemno) & "</FONT></td>"
'                Print #1, "<td width=15%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(rsTdayTran!STOCK_ORD) & "</FONT></td>"
'                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & SetSTOCKDESC(Null2String(rsTdayTran!STOCK_ORD)) & "</FONT></td>"
'                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(rsTdayTran!tranqty) & "</FONT></td>"
'                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(rsTdayTran!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
'                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
'                If knt <> 4 Then
'                    TOTALQTY = TOTALQTY + N2Str2IntZero(rsTdayTran!tranqty)
'                    TOTALPRICE = TOTALPRICE + N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE)
'                End If
'                Print #1, "</tr>"
'                If rsTdayTran.AbsolutePosition = MAX_ISS_LINE Then Exit Do
'                rsTdayTran.MoveNext
'            Loop
'            For cnt3 = 1 To cnt2
'                Print #1, "<tr>"
'                Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "</tr>"
'            Next
'            Print #1, "</table>"
'            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
'            If cntCOPY = 4 And knt < 3 Then
'                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'                Print #1, "<tr>"
'                Print #1, "<td width=10%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=15%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=35%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=8%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
'                Print #1, "</tr>"
'                Print #1, "</table>"
'            Else
'                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'                Print #1, "<tr>"
'                Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL RIV</FONT></td>"
'                Print #1, "<td align=right width=8%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & TOTALQTY & "</FONT></td>"
'                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'                Print #1, "<td align=right width=17%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & Format(TOTALPRICE, MAXIMUM_DIGIT) & "</FONT></td>"
'                Print #1, "</tr>"
'                Print #1, "</table>"
'            End If
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtPreparedBy.Text & "</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtIssuedBy.Text & "</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtApprovedBy.Text & "</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtRequestedBy.Text & "</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "<tr>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Requested By</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Approved By</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Issued By</FONT></td>"
'            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Received By</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
'            Print #1, "<tr>"
'            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'            Print #1, "</tr>"
'            Print #1, "</table>"
'            If knt <> 2 And knt <> 4 Then
'                Print #1, "<table>"
'                Print #1, "<tr>"
'                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'                Print #1, "</tr>"
'                Print #1, "</table>"
'                Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
'                Print #1, "<table>"
'                Print #1, "<tr>"
'                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
'                Print #1, "</tr>"
'                Print #1, "</table>"
'            End If
'        Next
'        Print #1, "</body></html>"
'        Close #1
'        On Error Resume Next
'        Open PMIS_REPORT_PATH & "ADB.HTML" For Input As #1
'        If EOF(1) Then
'            MsgSpeechBox "File Not Found!"
'            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
'        Else
'            Close #1
'            browRIV.Navigate "about:blank"
'            browRIV.Refresh
'            browRIV.Navigate PMIS_REPORT_PATH & "ADB.HTML"
'            DoEvents
'            If chkPreview.Value = 1 Then
'                browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
'            Else
'                browRIV.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
'            End If
'            Screen.MousePointer = 0
'        End If
'    End If
'    Set rsProfile = Nothing
'    Screen.MousePointer = 0
'End Sub
'
'Private Sub cmdPrintRIV_Click()
'    If rsOrd_Hd!TRANTYPE = "RIV" Then
'        SERVICEPISPRINTING
'        LogAudit "V", "PARTS RIV PRINTING"
'    End If
'    If rsOrd_Hd!TRANTYPE = "ADB" Then
'        ADBPRINTING
'        LogAudit "V", "PARTS ADB PRINTING"
'    End If
'
'    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", txtPreparedBy)
'    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", txtIssuedBy)
'    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", txtApprovedBy)
'
'    SendToBack
'End Sub
'
'Private Sub cmdTranCancel_Click()
'    SendToBack
'    StoreMemVars
'End Sub
'
'Private Sub cmdTranDelete_Click()
'
''updating code:    JAA - 07112007
'    On Error GoTo Errorcode:
'
'    Dim PnoOnhand, PnoTISSQTY, PnoIssuances            As Integer
'    If labDetID.Caption = "" Then
'        ShowNothingToDeleteMsg
'        Exit Sub
'    End If
'    If MsgQuestionBox("Delete This Parts, Are you Sure?", "Delete Parts Entry") = True Then
'        gconDMIS.Execute "delete from PMIS_TdayTran where id = " & labDetID.Caption
'        LogAudit "D", "CUSTOMER ORDER DETAIL", labDetID
'        ShowDeletedMsg
'    End If
'    'If txtTranType.Text <> "ADB" Then
'    '   Set rsPARTMAS = New ADODB.Recordset
'    '       rsPARTMAS.Open "select STOCKNO,onhand,TISSQTY,issuances from PMIS_PARTMAS where STOCKNO = '" & labPARTNO.Caption & "'", gconDMIS
'    '  If Not rsPARTMAS.EOF And Not rsPARTMAS.BOF Then
'    '     PnoOnhand = N2Str2IntZero(rsPARTMAS!Onhand)
'    '      PnoTISSQTY = N2Str2IntZero(rsPARTMAS!tissqty)
'    '      PnoIssuances = N2Str2IntZero(rsPARTMAS!issuances)
'    '      gconDMIS.Execute "update PMIS_PARTMAS set" & _
'           '                        " onhand = " & PnoOnhand + NumericVal(txtTranQty.Text) & "," & _
'           '                        " TISSQTY = " & PnoTISSQTY - NumericVal(txtTranQty.Text) & ", " & _
'           '                        " issuances = " & PnoIssuances - NumericVal(txtTranQty.Text) & _
'           '                        " where STOCKNO = " & N2Str2Null(rsPARTMAS!STOCKNO)
'    '   End If
'    'End If
'    Dim cnt                                            As Integer
'    Dim rsTdaytranDup                                  As ADODB.Recordset
'    Set rsTdaytranDup = New ADODB.Recordset
'    rsTdaytranDup.Open "select id,itemno from PMIS_TdayTran where [TYPE] = 'P' AND trantype = " & N2Str2Null(COUNTERTYPE) & " and tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " order by itemno asc", gconDMIS
'    If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
'        rsTdaytranDup.MoveFirst
'        cnt = 0
'        Do While Not rsTdaytranDup.EOF
'            cnt = cnt + 1
'            gconDMIS.Execute "update PMIS_TdayTran set itemno = " & Format(cnt, "0000") & " where id = " & rsTdaytranDup!ID
'            rsTdaytranDup.MoveNext
'        Loop
'    End If
'    FillDetails
'    gconDMIS.Execute "update PMIS_Ord_Hd set" & _
'                   " ttlinvamt = " & ORD_TOTUPRICE & "," & _
'                   " netinvamt = " & ORD_TOTINVAMT & _
'                   " where id = " & labID.Caption
'    rsRefresh
'    On Error Resume Next
'    rsOrd_Hd.Find "id = " & labID.Caption
'    cmdTranCancel.Value = True
'
'    Exit Sub
'Errorcode:
'    ShowVBError
'
'End Sub

Private Sub cmdTranSave_Click()
    Screen.MousePointer = 11
    On Error GoTo Errorcode

    If cboTranPartNo.Text = "" Then
        MsgSpeechBox "Warning: Part Number must have a value"
        On Error Resume Next
        cboTranPartNo.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Dim rsTDaytranClone                            As ADODB.Recordset
        Set rsTDaytranClone = New ADODB.Recordset
        rsTDaytranClone.Open "select trantype,tranno,itemno,STOCK_ORD from PMIS_TdayTran where [TYPE] = 'P' AND STOCK_ORD = '" & cboTranPartNo.Text & "' and trantype = '" & txtTranType.Text & "' and tranno =" & N2Str2Null(rsOrd_Hd!Tranno) & " order by itemno asc", gconDMIS
        If Not rsTDaytranClone.EOF And Not rsTDaytranClone.BOF Then
            MsgSpeechBox "Part Number already used in this transaction"
            On Error Resume Next
            cboTranPartNo.SetFocus
            Exit Sub
        End If
        Set rsTDaytranClone = Nothing
    End If

    Dim ORDTRANDATE, ORDTRANNO, ORDTRANTYPE            As String
    Dim ORDITEMNO, ORDSTOCK_ORD, ORDSTOCK_SUP          As String
    Dim ORDTRANQTY                                     As Integer
    Dim ORDUNIT                                        As String
    Dim ORDTRANUCOST                                   As Double
    Dim ORDSTATUS, ORDIN_OUT                           As String
    Dim ORDTRANINVAMT                                  As Double

    Dim CurONHAND, CurSAFESTOCK, CurTISSQTY            As Integer
    Dim curRESSERVICE, curIssuances, PrevCurOrdQty     As Integer

    If txtTranType.Text <> "ADB" Then
        Set rsPartMas = New ADODB.Recordset
        rsPartMas.Open "Select STOCKNO,onhand,sstock,resservice,TISSQTY,issuances from PMIS_PARTMAS where STOCKNO = '" & cboTranPartNo.Text & "' AND ACTIVE = 'Y'", gconDMIS
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then
            CurONHAND = N2Str2IntZero(rsPartMas!ONHAND)
            CurSAFESTOCK = N2Str2IntZero(rsPartMas!SSTOCK)
            CurTISSQTY = N2Str2IntZero(rsPartMas!TISSQTY)
            curRESSERVICE = N2Str2IntZero(rsPartMas!RESSERVICE)
            curIssuances = N2Str2IntZero(rsPartMas!issuances)
            If AddorEdit <> "ADD" Then
                PrevCurOrdQty = NumericVal(labPrevOrdQty.Caption)
                CurONHAND = CurONHAND + PrevCurOrdQty
                CurTISSQTY = CurTISSQTY - PrevCurOrdQty
                curIssuances = curIssuances - PrevCurOrdQty
            End If
            If CurONHAND <= 0 Then
                Screen.MousePointer = 0
                MsgSpeechBox "Out of Stock!"
                Exit Sub
            End If

            If txtTranType.Text = "CSH" Or txtTranType.Text = "CHG" Then
                If CurONHAND <= curRESSERVICE Then
                    Screen.MousePointer = 0
                    If MsgQuestionBox("Stock is Reserved for Service... Continue Anyway?", "Stock Status Alert!") = False Then
                        Exit Sub
                    End If
                End If
            End If

            If NumericVal(txtTranQty.Text) > CurONHAND Then
                Screen.MousePointer = 0
                MsgSpeechBox "Qty Ordered Exceeds Current Stock!"
                On Error Resume Next
                txtTranQty.SetFocus
                Exit Sub
            Else
                CurONHAND = CurONHAND - NumericVal(txtTranQty.Text)
            End If

            If CurONHAND < CurSAFESTOCK Then
                Screen.MousePointer = 0
                If MsgQuestionBox("Current On-hand is now below the Safety Stock Level... Proceed Anyway?", "Safety Stock Alert!") = False Then
                    Exit Sub
                End If
                Screen.MousePointer = 11
            End If
        Else
            Screen.MousePointer = 0
            MsgSpeechBox "Part Number Not Found!"
            Exit Sub
        End If
    End If

    ORDTRANDATE = N2Date2Null(txtTranDate.Text)
    ORDTRANTYPE = N2Str2Null(txtTranType.Text)
    ORDTRANNO = N2Str2Null(txtTranNo.Text)
    ORDITEMNO = N2Str2Null(Format(txtTranItemNo.Text, "0000"))
    ORDSTOCK_ORD = N2Str2Null(cboTranPartNo.Text)
    If txtTranType.Text = "ADB" Then ORDSTOCK_SUP = N2Str2Null(Left(txtTranDescription.Text, 100)) Else ORDSTOCK_SUP = N2Str2Null(cboTranPartNo.Text)
    ORDTRANQTY = NumericVal(txtTranQty.Text)
    ORDTRANUCOST = NumericVal(txtTranUCost.Text)
    ORDTRANINVAMT = NumericVal(txtTranUPrice.Text)
    If txtTranType.Text = "ADB" Then ORDIN_OUT = "'A'" Else ORDIN_OUT = "'O'"
    ORDSTATUS = "'N'"

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "insert into PMIS_TdayTran " & _
                         "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,tranuprice,lastupdate,usercode,status,in_out)" & _
                       " values ('P'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                       " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                       " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                       " " & ORDTRANUCOST & "," & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & ")"
    Else
        gconDMIS.Execute "update PMIS_TdayTran set" & _
                       " trandate = " & ORDTRANDATE & "," & _
                       " trantype = " & ORDTRANTYPE & "," & _
                       " tranno = " & ORDTRANNO & "," & _
                       " itemno = " & ORDITEMNO & "," & _
                       " STOCK_ORD = " & ORDSTOCK_ORD & "," & _
                       " STOCK_SUP = " & ORDSTOCK_SUP & "," & _
                       " tranqty = " & ORDTRANQTY & "," & _
                       " tranucost = " & ORDTRANUCOST & "," & _
                       " tranuprice = " & ORDTRANINVAMT & "," & _
                       " lastupdate = '" & LOGDATE & "'," & _
                       " status = " & ORDSTATUS & "," & _
                       " in_out = " & ORDIN_OUT & "," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "" & _
                       " where id = " & labDetID.Caption
    End If
    cleargrid grdDetails
    FillDetails
    gconDMIS.Execute "update PMIS_Ord_Hd set" & _
                   " totalqty = " & ORD_TOTQTY & "," & _
                   " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                   " netinvamt = " & ORD_TOTINVAMT & _
                   " where id = " & labID.Caption
    Dim rsPRS_Header                                   As ADODB.Recordset
    Dim rsPRS_Details                                  As ADODB.Recordset
    Set rsPRS_Header = New ADODB.Recordset
    Set rsPRS_Header = gconDMIS.Execute("Select * from PMIS_vw_PRS where REFPISNO = '" & cboRefPRSNo.Text & "'")
    If Not rsPRS_Header.EOF And Not rsPRS_Header.BOF Then
        Set rsPRS_Details = New ADODB.Recordset
        Set rsPRS_Details = gconDMIS.Execute("Select * from PMIS_vw_PRS_Tran Where Tranno = " & N2Str2Null(rsPRS_Header!Tranno) & " AND STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text))
        If Not rsPRS_Details.EOF And Not rsPRS_Details.BOF Then
            gconDMIS.Execute ("Update PMIS_vw_PRS_Tran set TRemarks = 'SERVED'  Where Tranno = " & N2Str2Null(rsPRS_Header!Tranno) & " AND STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text))
        Else
        End If
    Else
    End If
    'If txtTranType.Text <> "ADB" Then
    '   gconDMIS.Execute "update PMIS_PARTMAS set" & _
        '                     " onhand = " & CurONHAND & "," & _
        '                     " TISSQTY = " & CurTISSQTY + NumericVal(txtTranQty.Text) & ", " & _
        '                     " issuances = " & curIssuances + NumericVal(txtTranQty.Text) & _
        '                     " where STOCKNO = '" & cboTranPARTNO.Text & "'"
    'End If
    rsRefresh
    On Error Resume Next
    rsOrd_Hd.Find "id = " & labID.Caption
    StoreMemVars
    Screen.MousePointer = 0
    If AddorEdit = "ADD" Then cmdAddTran_Click Else cmdTranCancel.Value = True
    Exit Sub

Errorcode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdAdd_Click()

    If Function_Access(LOGID, "Acess_Add", LOCALACESS) = False Then Exit Sub
    AddorEdit = "ADD"
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    InitMemvars
    '=================================
    'updating code:     JAA - 12052007
    'To disable lstOrd_Hd Listview for Adding and Editing Transaction
    fraDetails.Enabled = False
    '=================================
    On Error Resume Next
    'cboChargeTo.SetFocus
    '    txtReferencePIS.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    '=================================
    'updating code:     JAA - 12052007
    'To enable lstOrd_Hd Listview
    fraDetails.Enabled = True
    '=================================
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", LOCALACESS) = False Then Exit Sub
    AddorEdit = "EDIT"
    PrevOrdType = txtTranType.Text
    PrevOrdNo = Format(txtTranNo.Text, "000000")
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    '=================================
    'updating code:     JAA - 12052007
    'To disable lstOrd_Hd Listview for Adding and Editing Transaction
    fraDetails.Enabled = False
    '=================================
    On Error Resume Next
    txtCustName.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    textSearch.SetFocus
End Sub

Sub FindDupTranno(DDD As String)
    On Error Resume Next
    rsOrd_Hd.Bookmark = rsFind(rsOrd_Hd.Clone, "tranno", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Private Sub cmdFirst_Click()
    rsOrd_Hd.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    rsOrd_Hd.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    rsOrd_Hd.MoveNext
    If rsOrd_Hd.EOF Then
        rsOrd_Hd.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsOrd_Hd.MovePrevious
    If rsOrd_Hd.BOF Then
        rsOrd_Hd.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim NextCunter                                     As String
    Dim rsCunter, rsfindDup                            As ADODB.Recordset
    Dim xSALES_ORIGIN, xSI_TYPE, xPAY_CLASS, xCHAR_YEAR, xCHAR_MONTH, xIS_SERIES, xTRACK_CODE As String

    Dim VcboSalesMan, VcboSMName, VTXTTranType         As String
    Dim VTXTTranNo, VTXTTranDate, VtxtCustCode         As String
    Dim VtxtCustName, VTXTChargeTo, VTXTREP_OR, VTXTREFPRSNO, VTXTRONO As String
    Dim VtxtTerms                                      As String
    Dim VTXTTTLInvAmt, VTXTDS1                         As Double
    Dim VTXTDS_Desc1                                   As String
    Dim VTXTDS_Amt1, VTXTNetInvAmt                     As Double
    Dim VTXTRemarks, VStatus, Vusercode                As String
    Dim VLastUpdate                                    As String
    Dim VIn_Process                                    As String
    Dim vtxtReferencePIS                               As String

    If txtTranType.Text <> "DR" Then
        If Trim(txtReferencePIS.Text) = "" Or Len(txtReferencePIS.Text) < 10 Then
            MsgBox "Invalid Reference PIS Number!", vbCritical, "PIS Required!"
            Exit Sub
        End If
    End If
    If Trim(txtTranType.Text) = "RIV" Then
        If Trim(txtRONO.Text) = "" Then
            MsgBox "RO Number is Required...", vbInformation, "Pls Input RO Number..."
            Exit Sub
        End If
    End If
    
    If IsNull(txtTranNo.Text) = True Then
        MsgSpeechBox "Transaction No. must not be empty"
        On Error Resume Next
        txtTranNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select trantype,tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Transaction No. already exist!"
                On Error Resume Next
                Exit Sub
            End If
        End If
    End If
    If txtTranDate.Text = "" Or IsDate(txtTranDate.Text) = False Then
        MsgSpeechBox "Invalid Transaction Date!"
        On Error Resume Next
        txtTranDate.SetFocus
        Exit Sub
    End If
    If txtTranType.Text = "CHG" Then
        If txtTerms.Text = "" Then
            MsgSpeechBox "Terms must have a value"
            On Error Resume Next
            txtTerms.SetFocus
            Exit Sub
        End If
    End If
    'If txtTranType.Text = "RIV" Or txtTranType.Text = "ADB" Then
    '   VcboSalesMan = "NULL"
    '   VcboSMName = "NULL"
    'Else
    VcboSalesMan = N2Str2Null(cboSalesMan.Text)
    VcboSMName = N2Str2Null(cboSMName.Text)
    'End If

    NextCunter = NumericVal(txtTranNo.Text) + 1

    VTXTTranType = N2Str2Null(txtTranType.Text)
    VTXTTranNo = N2Str2Null(txtTranNo.Text)
    VTXTTranDate = N2Date2Null(txtTranDate.Text)
    VtxtCustCode = N2Str2Null(txtCustCode.Text)
    VtxtCustName = N2Str2Null(txtCustName.Text)
    vtxtReferencePIS = N2Str2Null(txtReferencePIS.Text)
    VTXTREFPRSNO = N2Str2Null(cboRefPRSNo.Text)
    VIn_Process = "'Y'"
    'If cboChargeTo.Text) = "MECHANICAL" Then
    '   VTXTChargeTo = "'MEC'"
    'ElseIf cboChargeTo.Text) = "COMPANY" Then
    '   VTXTChargeTo = "'COM'"
    'ElseIf cboChargeTo.Text) = "WARRANTY" Then
    '   VTXTChargeTo = "'WAR'"
    'ElseIf cboChargeTo.Text) = "TINSMITH" Then
    '   VTXTChargeTo = "'TIN'"
    'ElseIf cboChargeTo.Text) = "FLEET" Then
    '   VTXTChargeTo = "'FLE'"
    'ElseIf cboChargeTo.Text) = "VARIOUS" Then
    VTXTChargeTo = "'VAR'"
    'ElseIf cboChargeTo.Text) = "PARTS CLAIM" Then
    '   VTXTChargeTo = "'PCL'"
    'Else
    '   MsgSpeechBox "Invalid Issuance Charge..."
    '   On Error Resume Next
    'cboChargeTo.SetFocus
    '   Exit Sub
    'End If

    Dim RRTRANDATE, RRTRANNO, RRTRANTYPE               As String
    Dim RRITEMNO, RRSTOCK_ORD, RRSTOCK_SUP             As String
    Dim RRTRANQTY                                      As Integer
    Dim RRTRANUCOST, RRTRANINVAMT                      As Double
    Dim RRIN_OUT, RRSTATUS                             As String

    'VTXTChargeTo = N2Str2Null(txtChargeTo.Text)
    VTXTRONO = N2Str2Null(txtRONO.Text)
    If Len(txtRONO.Text) = 7 Then
        VTXTREP_OR = "'" & Left(txtRONO.Text, 1) & "-" & Right(txtRONO.Text, 6) & "'"
    Else
        VTXTREP_OR = "NULL"
    End If
    VtxtTerms = N2Str2Null(txtTerms.Text)
    VTXTTTLInvAmt = NumericVal(txtTTLInvAmt.Text)
    VTXTDS1 = NumericVal(txtDS1.Text)
    VTXTDS_Desc1 = N2Str2Null(txtDS_Desc1.Text)
    VTXTDS_Amt1 = NumericVal(txtDS_Amt1.Text)
    VTXTNetInvAmt = NumericVal(txtNetInvAmt.Text)
    If txtRemarks.Text = "Pls Type Your Message Here!" Then VTXTRemarks = "NULL" Else VTXTRemarks = N2Str2Null(Trim(txtRemarks.Text))
    VStatus = "'N'"
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"

    xSALES_ORIGIN = N2Str2Null(Mid(txtReferencePIS, 3, 1))
    xSI_TYPE = N2Str2Null(Mid(txtReferencePIS, 4, 1))
    xPAY_CLASS = N2Str2Null(Mid(txtReferencePIS, 5, 1))
    xCHAR_YEAR = N2Str2Null(Mid(txtReferencePIS, 6, 2))
    xCHAR_MONTH = N2Str2Null(Mid(txtReferencePIS, 8, 1))
    xIS_SERIES = N2Str2Null(Mid(txtReferencePIS, 9, 3))
    xTRACK_CODE = N2Str2Null(Mid(txtReferencePIS, 12, 1))

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into PMIS_Ord_Hd" & _
                       " (TYPE,trantype,tranno,trandate,custcode,custname,chargeto,REFPRSNO,rono,rep_or,salesman,smname,terms,ttlinvamt,ds1,ds_desc1,ds_amt1,netinvamt,remarks,status,usercode,lastupdate,In_Process,REFPISNO,SALES_ORIGIN,SI_TYPE,PAY_CLASS,CHAR_YEAR,CHAR_MONTH,IS_SERIES,TRACK_CODE)" & _
                       " values ('P'," & VTXTTranType & ", " & VTXTTranNo & ", " & VTXTTranDate & ", " & _
                       " " & VtxtCustCode & ", " & VtxtCustName & ", " & VTXTChargeTo & "," & VTXTREFPRSNO & _
                         ", " & VTXTRONO & "," & VTXTREP_OR & ", " & VcboSalesMan & ", " & VcboSMName & _
                         ", " & VtxtTerms & ", " & VTXTTTLInvAmt & _
                         ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                         ", " & VTXTNetInvAmt & ", " & VTXTRemarks & _
                         ", " & VStatus & ", " & Vusercode & ", " & VLastUpdate & "," & VIn_Process & "," & vtxtReferencePIS & ", " & xSALES_ORIGIN & ", " & xSI_TYPE & ", " & xPAY_CLASS & ", " & xCHAR_YEAR & ", " & xCHAR_MONTH & ", " & xIS_SERIES & ", " & xTRACK_CODE & ")"
        LogAudit "A", "PARTS CUSTOMER ORDER", txtTranNo & txtCustCode
    Else

        gconDMIS.Execute "update PMIS_Ord_Hd set" & _
                       " trantype = " & VTXTTranType & "," & _
                       " tranno = " & VTXTTranNo & "," & _
                       " trandate = " & VTXTTranDate & "," & _
                       " custcode = " & VtxtCustCode & "," & _
                       " custname = " & VtxtCustName & "," & _
                       " chargeto = " & VTXTChargeTo & "," & _
                       " REFPRSNO = " & VTXTREFPRSNO & "," & _
                       " rono = " & VTXTRONO & "," & _
                       " rep_or = " & VTXTREP_OR & "," & _
                       " salesman = " & VcboSalesMan & "," & _
                       " smname = " & VcboSMName & "," & _
                       " terms = " & VtxtTerms & "," & _
                       " ttlinvamt = " & VTXTTTLInvAmt & "," & _
                       " ds1 = " & VTXTDS1 & "," & _
                       " ds_desc1 = " & VTXTDS_Desc1 & "," & _
                       " ds_amt1 = " & VTXTDS_Amt1 & "," & _
                       " netinvamt = " & VTXTNetInvAmt & "," & _
                       " remarks = " & VTXTRemarks & ", " & _
                       " status = " & VStatus & ", " & _
                       " usercode = " & Vusercode & ", " & _
                       " In_Process = " & VIn_Process & ", " & _
                       " REFPISNO = " & vtxtReferencePIS & ", " & _
                       " lastupdate = " & VLastUpdate & _
                       " where id = " & labID.Caption

        gconDMIS.Execute "update PMIS_Ord_Hd set" & _
                       " SALES_ORIGIN = " & xSALES_ORIGIN & "," & _
                       " SI_TYPE = " & xSI_TYPE & "," & _
                       " PAY_CLASS = " & xPAY_CLASS & "," & _
                       " CHAR_YEAR = " & xCHAR_YEAR & "," & _
                       " CHAR_MONTH = " & xCHAR_MONTH & "," & _
                       " IS_SERIES = " & xIS_SERIES & "," & _
                       " TRACK_CODE = " & xTRACK_CODE & "" & _
                       " where id = " & labID.Caption

        gconDMIS.Execute "update PMIS_TdayTran set" & _
                       " trantype = " & VTXTTranType & "," & _
                       " trandate = " & VTXTTranDate & "," & _
                       " tranno = " & VTXTTranNo & _
                       " where [TYPE] = 'P' AND trantype = '" & PrevOrdType & "' and tranno = '" & PrevOrdNo & "'"
        LogAudit "E", "PARTS CUSTOMER ORDER", txtTranNo & txtCustCode
    End If
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "update PMIS_Counter set nextnumber = '" & NextCunter & "', lastupdate = '" & LOGDATE & "', usercode = '" & "USER" & "' where [TYPE] = 'P' AND modul = " & VTXTTranType
        Call FillGrid
    Else
        rsRefresh
        rsOrd_Hd.Find "Tranno = " & VTXTTranNo
        cmdCancel.Value = True
        cleargrid grdDetails
        FillDetails
        gconDMIS.Execute "update PMIS_Ord_Hd set" & _
                       " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                       " netinvamt = " & ORD_TOTINVAMT & _
                       " where [TYPE] = 'P' AND tranno = " & VTXTTranNo & " and trantype = " & VTXTTranType
    End If
    '=================================
    'updating code:     JAA - 12052007
    'To enable lstOrd_Hd Listview
    fraDetails.Enabled = True
    '=================================
    rsRefresh
    rsOrd_Hd.Find "tranno = " & VTXTTranNo
    cmdCancel.Value = True
    On Error GoTo Errorcode
    If AddorEdit = "ADD" Then
        Dim rsTdaytranDup, rstdaytranDUp2              As ADODB.Recordset
        Dim rsPRS_HD As ADODB.Recordset
        Dim varPmasTrecqty, varPmasOnOrder, varPmasOnhand As Long
        Dim rsPartMasClone                             As ADODB.Recordset
        Dim Iss_Cnt As Integer
        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "select trantype,tranno from PMIS_TdayTran where [TYPE] = 'P' AND trantype = '" & COUNTERTYPE & "' and tranno = " & N2Str2Null(rsOrd_Hd!Tranno), gconDMIS
        If rsTdaytranDup.EOF And rsTdaytranDup.BOF Then
            rsTdaytranDup.Close
            Set rsPRS_HD = New ADODB.Recordset
            Set rsPRS_HD = gconDMIS.Execute("Select * from PMIS_vw_PRS where refpisno = '" & cboRefPRSNo.Text & "'")
            If Not rsPRS_HD.EOF And Not rsPRS_HD.BOF Then
                Set rstdaytranDUp2 = New ADODB.Recordset
                'rstdaytranDUp2.Open "select trantype,tranno,STOCK_ORD,STOCK_SUP,itemno,tranqty,traninvamt,tranucost from PMIS_TdayTran where trantype = 'PO' and tranno = " & N2Str2Null(rsRR_HD!PONO), gconDMIS
                '==================================================================================================================================================================
                'REM - UDPDATED FML - 05042007
                rstdaytranDUp2.Open "select id,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost,tranuprice from PMIS_TdayTran where trantype = 'PRS' and tranno = " & N2Str2Null(rsPRS_HD!Tranno) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                'rstdaytranDUp2.Open "select id,itemno,STOCK_ORD,STOCK_SUP,QTY_ALLOCATED AS tranqty,traninvamt,tranucost from PMIS_vw_ConfirmedPO where PO_NO = " & N2Str2Null(txtPONo.Text) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                '==================================================================================================================================================================
                If Not rstdaytranDUp2.EOF And Not rstdaytranDUp2.BOF Then
                    rstdaytranDUp2.MoveFirst: Iss_Cnt = 0
                    Do While Not rstdaytranDUp2.EOF
                        Set rsPartMasClone = New ADODB.Recordset
                        Set rsPartMasClone = gconDMIS.Execute("Select STOCKNO,ONHAND from PMIS_StockMas where TYPE = 'P' and STOCKNO = " & N2Str2Null(rstdaytranDUp2!STOCK_ORD))
                        If Not rsPartMasClone.EOF And Not rsPartMasClone.BOF Then
                           If N2Str2Zero(rsPartMasClone!ONHAND) > 0 Then
                                Iss_Cnt = Iss_Cnt + 1
                                RRTRANDATE = N2Str2Null(rsOrd_Hd!trandate)
                                RRTRANTYPE = "'" & COUNTERTYPE & "'"
                                RRTRANNO = N2Str2Null(rsOrd_Hd!Tranno)
                                RRITEMNO = N2Str2Null(Format(Null2String(rstdaytranDUp2!itemno), "0000"))
                                RRSTOCK_ORD = N2Str2Null(rstdaytranDUp2!STOCK_ORD)
                                RRSTOCK_SUP = N2Str2Null(rstdaytranDUp2!STOCK_SUP)
                                RRTRANQTY = N2Str2IntZero(rstdaytranDUp2!tranqty)
                                RRTRANINVAMT = N2Str2Zero(rstdaytranDUp2!TRANUPRICE)
                                RRTRANUCOST = N2Str2Zero(rstdaytranDUp2!TRANUPRICE)
                                RRIN_OUT = "'O'"
                                RRSTATUS = "'N'"
            
                                gconDMIS.Execute "insert into PMIS_TdayTran " & _
                                                 "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice,lastupdate,usercode,status,in_out)" & _
                                               " values ('P'," & RRTRANDATE & ", '" & COUNTERTYPE & "', " & RRTRANNO & "," & _
                                               " " & RRITEMNO & "," & RRSTOCK_ORD & "," & _
                                               " " & RRSTOCK_SUP & ", " & RRTRANQTY & "," & _
                                               " " & RRTRANUCOST & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & ")"
                            Else
                                MsgBox "Requested Part No. " & Null2String(rstdaytranDUp2!STOCK_ORD) & " doesn't have Stock in your Master File", vbInformation, "Cannot Add Parts!"
                            End If
                        Else
                            MsgBox "Requested Part No. " & Null2String(rstdaytranDUp2!STOCK_ORD) & " is not yet active in your Master File", vbInformation, "Cannot Add Parts!"
                        End If
                        rstdaytranDUp2.MoveNext
                    Loop
                End If
            End If
            cleargrid grdDetails
            FillDetails
            'cmdAddTran_Click
        Else
            cleargrid grdDetails
            FillDetails
            cmdAddTran_Click
        End If
    End If
    
    '====
    If AddorEdit = "ADD" Then
        InsertAdvanceBill
        'cmdAddTran_Click
    End If
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Command1_Click()
'frmAllCustomer.Show
    frmCustomerSearch.Show 1
End Sub

Private Sub Command2_Click()
    'txtTranUPrice.Enabled = True
    On Error Resume Next
    'txtTranUPrice.Enabled = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim fild                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    fild = grdDetails.Text
    Select Case KeyCode
        Case vbKeyEscape
            If Picture1.Visible = True Then
                SendToBack
                StoreMemVars
            End If
        'Case vbKeyF1
        '    If Picture1.Visible = False Then Command2.Value = True
        'Case vbKeyF2
        '    If Command1.Visible = True And Command1.Enabled = True Then Command1.Value = True
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(rsOrd_Hd!Status) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change..."
                ElseIf Null2String(rsOrd_Hd!Status) = "B" Then
                    MsgSpeechBox "Transactions are Already Billed-Out and cannot be Change..."
                ElseIf Null2String(rsOrd_Hd!Status) = "P" Then
                    MsgSpeechBox "Transactions are Already Posted and cannot be Change..."
                Else
                    cmdAddTran_Click
                End If
            End If
        Case vbKeyF4
            If fild <> "" And fild <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(rsOrd_Hd!Status) <> "P" And Null2String(rsOrd_Hd!Status) <> "C" And Null2String(rsOrd_Hd!Status) <> "B" Then
                        grdDetails_DblClick
                    End If
                End If
            End If
        Case vbKeyF5
            If fild <> "" And fild <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(rsOrd_Hd!Status) <> "P" And Null2String(rsOrd_Hd!Status) <> "C" And Null2String(rsOrd_Hd!Status) <> "B" Then
                        grdDetails_DblClick
                        'cmdTranDelete_Click
                    End If
                End If
            End If
        Case vbKeyF8
            If cmdPost.Enabled = True Then cmdPost.Value = True
        Case vbKeyF12
            If Picture1.Visible = True Then
                If Function_Access(LOGID, "Acess_UNPost", LOCALACESS) = False Then Exit Sub
                If Null2String(rsOrd_Hd!Status) = "P" Then
                    'If LOGLEVEL <> "ADM" Then
                    '   MsgBox "Warning: Your account is not allowed to UnPost this transaction!", vbCritical, "Error"
                    '   Exit Sub
                    'End If
                    If MsgQuestionBox("Are you sure you want to UnPost this Transaction?", "UnPost Transaction") = True Then
                        Dim PCurOnHand, PCurTISSQTY, PCurIssuances As Integer
                        Dim rsTdaytranDup, rsPartmasDup As ADODB.Recordset

                        Set rsTdaytranDup = New ADODB.Recordset
                        rsTdaytranDup.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_TdayTran where [TYPE] = 'P' AND tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " and trantype = " & N2Str2Null(rsOrd_Hd!TRANTYPE), gconDMIS, adOpenForwardOnly, adLockReadOnly
                        If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
                            rsTdaytranDup.MoveFirst
                            Do While Not rsTdaytranDup.EOF
                                Set rsPartmasDup = New ADODB.Recordset
                                rsPartmasDup.Open "select STOCKNO,onhand,tissqty,TISSQTY,issuances,REQSERVED,S_REQSERVED from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD) & " AND ACTIVE = 'Y'", gconDMIS
                                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                                    PCurOnHand = N2Str2IntZero(rsPartmasDup!ONHAND) + N2Str2Zero(rsTdaytranDup!tranqty)
                                    PCurTISSQTY = N2Str2IntZero(rsPartmasDup!TISSQTY) - N2Str2Zero(rsTdaytranDup!tranqty)
                                    PCurIssuances = N2Str2IntZero(rsPartmasDup!issuances) - N2Str2Zero(rsTdaytranDup!tranqty)
                                    If Null2String(rsOrd_Hd!SALES_ORIGIN) = "W" Then
                                        gconDMIS.Execute "update PMIS_PARTMAS set" & _
                                                       " REQSERVED = " & N2Str2IntZero(rsPartmasDup!REQServed) - N2Str2Zero(rsTdaytranDup!tranqty) & _
                                                       " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                                    Else
                                        gconDMIS.Execute "update PMIS_PARTMAS set" & _
                                                       " S_REQSERVED = " & N2Str2IntZero(rsPartmasDup!S_REQServed) - N2Str2Zero(rsTdaytranDup!tranqty) & _
                                                       " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                                    End If
                                    gconDMIS.Execute "update PMIS_PARTMAS set" & _
                                                   " onhand = " & PCurOnHand & "," & _
                                                   " tissqty = " & PCurTISSQTY & "," & _
                                                   " issuances = " & PCurIssuances & "," & _
                                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                   " lastupdate = '" & LOGDATE & "'" & _
                                                   " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                                    gconDMIS.Execute "update PMIS_TdayTran set" & _
                                                   " status = 'N'," & _
                                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                   " lastupdate = '" & LOGDATE & "'" & _
                                                   " where id = " & rsTdaytranDup!ID
                                End If
                                rsTdaytranDup.MoveNext
                            Loop
                        End If
                        gconDMIS.Execute "update PMIS_Ord_Hd set" & _
                                       " status = 'N'," & _
                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                       " lastupdate = '" & LOGDATE & "'" & _
                                       " where id = " & labID.Caption
                        rsRefresh
                        On Error Resume Next
                        rsOrd_Hd.Find "id =" & labID.Caption
                        StoreMemVars
                    End If
                    Set rsTdaytranDup = Nothing
                    Set rsPartmasDup = Nothing
                End If
                LogAudit "U", "CUSTOMER ORDER", txtTranNo & txtCustCode
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    LOCALACESS = ""
    
    If COUNTERTYPE = "RIV" Then
        LOCALACESS = "PARTS ISSUANCE SERVICE ISSUANCE"
    ElseIf COUNTERTYPE = "ADB" Then
        LOCALACESS = "PARTS ADVANCE BILL DATA ENTRY"
    ElseIf COUNTERTYPE = "CSH" Then
        LOCALACESS = "PARTS ISSUANCE COUNTER CASH"
    ElseIf COUNTERTYPE = "CHG" Then
        LOCALACESS = "PARTS ISSUANCE COUNTER CHARGE"
    ElseIf COUNTERTYPE = "DR" Then
        LOCALACESS = "PARTS DR OUT ISSUANCE"
    End If
    
    
    CenterMe frmMain, Me, 1: PMIS_ORDER_SHOW = True
    textSearch.Text = "":    'Picture5.ZOrder 0
    If COUNTERTYPE <> "RIV" And COUNTERTYPE <> "ADB" Then
        Command1.Visible = True
        Command1.Enabled = True
        optRONo.Enabled = False
    Else
        Command1.Enabled = False
        Command1.Visible = False
    End If
    If COUNTERTYPE = "DR" Then cmdPISNum.Enabled = False
    If COUNTERTYPE = "CSH" Then optCASH.Value = True
    If COUNTERTYPE = "CHG" Then optCHARGE.Value = True
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    
    
    
    InitMemvars
    If LOGLEVEL = "ADM" Then
        'txtTranUPrice.Enabled = True
    Else
        If COUNTERTYPE = "ADB" Then
            'txtTranUPrice.Enabled = True
        Else
            txtTranUPrice.Enabled = False
        End If
    End If
    rsRefresh
    On Error Resume Next
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then rsOrd_Hd.MoveLast
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    If LOGLEVEL = "RIV USER" Then
        If COUNTERTYPE = "ADB" Then
            Me.Caption = "ADVANCE BILL DATA ENTRY"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = 'ADB' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If COUNTERTYPE = "RIV" Then
            Me.Caption = "Requisition Issuance Data Entry"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = 'RIV' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        InitCboChargeToWarehouse
    Else
        'If LOGLEVEL = "SUPERVISOR" Or LOGLEVEL = "MANAGER" Or LOGLEVEL = "AUTHOR" Or LOGLEVEL = "ADM" Then
        If COUNTERTYPE = "CSH" Then
            Me.Caption = "Parts Issuance Slip Data Entry (Over the Counter)"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = 'CSH' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If COUNTERTYPE = "CHG" Then
            Me.Caption = "Charge Counter Issuance Data Entry"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = 'CHG' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If COUNTERTYPE = "RIV" Then
            Me.Caption = "Parts Issuance Slip Data Entry (Service Requisition)"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = 'RIV' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If COUNTERTYPE = "DR" Then
            Me.Caption = "DR Out Issuance Data Entry"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = 'DR' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If COUNTERTYPE = "ADB" Then
            Me.Caption = "Advance Bill Data Entry"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = 'ADB' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        InitCboChargeToCounter
    End If
End Sub

Sub InitCboChargeToWarehouse()
    cboChargeTo.Clear
    cboChargeTo.AddItem "MECHANICAL"
    cboChargeTo.AddItem "COMPANY"
    cboChargeTo.AddItem "WARRANTY"
    cboChargeTo.AddItem "TINSMITH"
    cboChargeTo.AddItem "VARIOUS"
    cboChargeTo.AddItem "FLEET"
    cboChargeTo.AddItem "PARTS CLAIM"
    cboChargeTo.Text = "MECHANICAL"
End Sub

Sub InitCboChargeToCounter()
    cboChargeTo.Clear
    cboChargeTo.AddItem "MECHANICAL"
    cboChargeTo.AddItem "COMPANY"
    cboChargeTo.AddItem "WARRANTY"
    cboChargeTo.AddItem "TINSMITH"
    cboChargeTo.AddItem "VARIOUS"
    cboChargeTo.AddItem "FLEET"
    cboChargeTo.AddItem "PARTS CLAIM"
    cboChargeTo.Text = "VARIOUS"
End Sub

Sub InitMemvars()
    If COUNTERTYPE = "RIV" Then
        Set rsCunter = New ADODB.Recordset
        rsCunter.Open "select * from PMIS_Counter where [TYPE] = 'P' AND modul = 'RIV'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            txtTranNo.Text = Format(Null2String(rsCunter!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtRONO.Enabled = True
        txtTerms.Enabled = False
    End If
    If COUNTERTYPE = "CSH" Then
        Set rsCunter = New ADODB.Recordset
        rsCunter.Open "select * from PMIS_Counter where [TYPE] = 'P' AND modul = 'CSH'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            txtTranNo.Text = Format(Null2String(rsCunter!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtRONO.Enabled = False
        txtTerms.Enabled = False
    End If
    If COUNTERTYPE = "CHG" Then
        Set rsCunter = New ADODB.Recordset
        rsCunter.Open "select * from PMIS_Counter where [TYPE] = 'P' AND modul = 'CHG'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            txtTranNo.Text = Format(Null2String(rsCunter!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtRONO.Enabled = False
        txtTerms.Enabled = True
    End If
    If COUNTERTYPE = "DR" Then
        Set rsCunter = New ADODB.Recordset
        rsCunter.Open "select * from PMIS_Counter where [TYPE] = 'P' AND modul = 'DR'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            txtTranNo.Text = Format(Null2String(rsCunter!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtRONO.Enabled = False
        txtTerms.Enabled = True
    End If
    If COUNTERTYPE = "ADB" Then
        Set rsCunter = New ADODB.Recordset
        rsCunter.Open "select * from PMIS_Counter where [TYPE] = 'P' AND modul = 'ADB'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            txtTranNo.Text = Format(Null2String(rsCunter!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtRONO.Enabled = True
        txtTerms.Enabled = False
    End If
    txtTranDate.Text = LOGDATE
    txtCustCode.Text = ""
    txtCustName.Text = ""
    txtChargeTo.Text = "VAR"
    txtReferencePIS.Text = ""
    cboRefPRSNo.Clear
    txtRONO.Text = ""
    txtTerms.Text = ""
    txtTTLInvAmt.Text = "0.00"
    txtDS1.Text = "0"
    txtDS_Desc1.Text = "0.00"
    txtDS_Amt1.Text = "0.00"
    txtNetInvAmt.Text = "0.00"
    txtRemarks.Text = "Pls Type Your Message Here!"
    labPosted.Caption = ""
    InitCbo
    InitGrid
    cleargrid grdDetails
    SendToBack
    InitSignatories
End Sub

Sub InitSignatories()
    txtPreparedBy.Text = ""
    txtIssuedBy.Text = ""
    txtRequestedBy.Text = ""
    txtApprovedBy.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        labID.Caption = rsOrd_Hd!ID
        txtTranType.Text = Null2String(rsOrd_Hd!TRANTYPE)
        'If txtTranType.Text = "RIV" Then
        cboSMName.Enabled = True
        txtTranNo.Text = Null2String(rsOrd_Hd!Tranno)
        txtTranDate.Text = Null2String(rsOrd_Hd!trandate)
        txtCustCode.Text = Null2String(rsOrd_Hd!custcode)
        txtCustName.Text = Null2String(rsOrd_Hd!custname)
        txtReferencePIS.Text = Null2String(rsOrd_Hd!refpisno)
        cboRefPRSNo.Text = Null2String(rsOrd_Hd!refpRsno)
        
        '========================================
        'UPDATING CODE:     JAA - 11072007
        If Mid(txtReferencePIS, 5, 1) = "W" Then
            txtTranUPrice.Enabled = True
        Else
            txtTranUPrice.Enabled = False
        End If
        '========================================
        
        If Null2String(rsOrd_Hd!chargeto) = "MEC" Then
            cboChargeTo.Text = "MECHANICAL"
        ElseIf Null2String(rsOrd_Hd!chargeto) = "COM" Then
            cboChargeTo.Text = "COMPANY"
        ElseIf Null2String(rsOrd_Hd!chargeto) = "WAR" Then
            cboChargeTo.Text = "WARRANTY"
        ElseIf Null2String(rsOrd_Hd!chargeto) = "TIN" Then
            cboChargeTo.Text = "TINSMITH"
        ElseIf Null2String(rsOrd_Hd!chargeto) = "FLE" Then
            cboChargeTo.Text = "FLEET"
        ElseIf Null2String(rsOrd_Hd!chargeto) = "VAR" Then
            cboChargeTo.Text = "VARIOUS"
        ElseIf Null2String(rsOrd_Hd!chargeto) = "PCL" Then
            cboChargeTo.Text = "PARTS CLAIM"
        Else
            cboChargeTo.Text = ""
        End If
        txtRONO.Text = Null2String(rsOrd_Hd!rono)
        cboSMName.Text = FillSalesMan(Null2String(rsOrd_Hd!salesman))
        txtTerms.Text = Null2String(rsOrd_Hd!Terms)
        txtTTLInvAmt.Text = ToDoubleNumber(N2Str2Zero(rsOrd_Hd!ttlinvamt))
        txtDS1.Text = N2Str2IntZero(rsOrd_Hd!ds1)
        txtDS_Desc1.Text = Null2String(rsOrd_Hd!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(rsOrd_Hd!ds_amt1))
        txtNetInvAmt.Text = ToDoubleNumber(N2Str2Zero(rsOrd_Hd!netinvamt))
        txtRemarks.Text = Null2String(rsOrd_Hd!remarks)
        If CheckIfROBilled(Null2String(rsOrd_Hd!rono)) = True Then
            labPosted.Caption = "BILLED OUT"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = False
        Else
            If Null2String(rsOrd_Hd!Status) = "C" Then
                labPosted.Caption = "CANCELLED"
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = False
            ElseIf Null2String(rsOrd_Hd!Status) = "P" Then
                labPosted.Caption = "POSTED"
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = True
            Else
                labPosted.Caption = ""
                cmdEdit.Enabled = True
                If LOGLEVEL = "ADM" Then cmdCancelCO.Enabled = True
                cmdPost.Enabled = True
                cmdPrint.Enabled = False
            End If
            If Null2String(rsOrd_Hd!In_Process) = "N" Then
                labPosted.Caption = "RELEASED"
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = False
            End If
        End If
        cleargrid grdDetails
        FillDetails
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Function CheckIfROBilled(XXX As String) As Boolean
Dim rsRo_det As ADODB.Recordset
Set rsRo_det = New ADODB.Recordset
Set rsRo_det = gconDMIS.Execute("Select INVOICE from CSMS_REPOR where INVOICE IS NOT NULL AND REP_OR = " & N2Str2Null(XXX))
If Not rsRo_det.EOF And Not rsRo_det.BOF Then
   CheckIfROBilled = True
Else
   CheckIfROBilled = False
End If
Set rsRo_det = Nothing
End Function

Sub InitGrid()
    With grdDetails
        .Rows = 7
        .ColWidth(0) = 1
        .ColWidth(1) = 1000
        .ColWidth(2) = 1500
        .ColWidth(3) = 2200
        .ColWidth(4) = 1000
        .ColWidth(5) = 1200
        .ColWidth(6) = 1300
        .Row = 0
        .Col = 1
        .Text = "Item"
        .Col = 2
        .Text = "Part Number"
        .Col = 3
        .Text = "Description"
        .Col = 4
        .Text = "QTY"
        .Col = 5
        .Text = "Price"
        .Col = 6
        .Text = "Extend Price"
    End With
End Sub

Sub FillDetails()
    kcnt = 0: ORD_TOTUPRICE = 0: ORD_TOTINVAMT = 0: ORD_TOTVAT = 0: ORD_TOTQTY = 0
    Dim STOCKDESCription                               As String
    Set rsTdayTran = New ADODB.Recordset
    'rsTdaytran.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " and trantype = " & N2Str2Null(rsOrd_Hd!trantype) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    rsTdayTran.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where [TYPE] = 'P' AND tranno = " & N2Str2Null(txtTranNo.Text) & " and trantype = " & N2Str2Null(txtTranType.Text) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        cboChargeTo.Enabled = False
        Screen.MousePointer = 11
        rsTdayTran.MoveFirst
        Do While Not rsTdayTran.EOF
            kcnt = kcnt + 1
            If txtTranType.Text = "ADB" Then
                STOCKDESCription = Null2String(rsTdayTran!STOCK_SUP)
            Else
                STOCKDESCription = SetSTOCKDESC(Null2String(rsTdayTran!STOCK_SUP))
            End If
            grdDetails.AddItem rsTdayTran!ID & Chr(9) & Null2String(rsTdayTran!itemno) & Chr(9) & _
                               Null2String(rsTdayTran!STOCK_ORD) & Chr(9) & _
                               STOCKDESCription & Chr(9) & _
                               N2Str2IntZero(rsTdayTran!tranqty) & Chr(9) & _
                               Format(N2Str2Zero(rsTdayTran!TRANUPRICE), MAXIMUM_DIGIT) & Chr(9) & _
                               Format(N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE), MAXIMUM_DIGIT)
            ORD_TOTQTY = ORD_TOTQTY + N2Str2IntZero(rsTdayTran!tranqty)
            ORD_TOTUPRICE = ORD_TOTUPRICE + (N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE))
            ORD_TOTINVAMT = ORD_TOTINVAMT + (N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE))
            rsTdayTran.MoveNext
        Loop
        If NumericVal(txtDS1.Text) <> 0 Then
            If txtDS_Desc1.Text = "" Then
                txtDS_Desc1.Text = "DISCOUNT"
            End If
            txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
            txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
        Else
            txtDS_Desc1.Text = ""
            txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
            txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
        End If
        ORD_TOTINVAMT = ORD_TOTINVAMT - NumericVal(txtDS_Amt1.Text)
        If kcnt <> 0 Then grdDetails.RemoveItem 1
        Screen.MousePointer = 0
    Else
        cboChargeTo.Enabled = True
        cleargrid grdDetails
    End If
End Sub

Function FillSalesMan(XXX As String) As String
    Set rsSalesMan = New ADODB.Recordset
    rsSalesMan.Open "select empno,signname from PMIS_vw_SalesMan where empno = '" & XXX & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSalesMan.EOF And Not rsSalesMan.BOF Then
        FillSalesMan = Null2String(rsSalesMan!signname)
        cboSalesMan.Text = Null2String(rsSalesMan!empno)
    Else
        cboSalesMan.Text = ""
    End If
End Function

Sub InitCbo()
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "select id,STOCKNO,STOCKDESC from PMIS_PARTMAS where ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        rsPartMas.MoveFirst
        cboTranPartNo.Clear
        Do While Not rsPartMas.EOF
            cboTranPartNo.AddItem Null2String(rsPartMas!STOCKNO)
            rsPartMas.MoveNext
        Loop
    End If
    FillCboSalesMan
End Sub

Sub FillCboSalesMan()
    Set rsSalesMan = New ADODB.Recordset
    rsSalesMan.Open "select empno,signname from PMIS_vw_SalesMan order by signname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSalesMan.EOF And Not rsSalesMan.BOF Then
        rsSalesMan.MoveFirst: cboSalesMan.Clear: cboSMName.Clear
        Do While Not rsSalesMan.EOF
            cboSalesMan.AddItem Null2String(rsSalesMan!empno)
            cboSMName.AddItem Null2String(rsSalesMan!signname)
            rsSalesMan.MoveNext
        Loop
    Else
        cboSalesMan.Clear: cboSMName.Clear
    End If
End Sub

Sub SetCustInfo(rep As String)
    rep = Left(rep, 1) & "-" & Right(rep, 6)
    Set rsREPOR = New ADODB.Recordset
    'rsREPOR.Open "select rep_or,niym,acct_no,invoice from CSMS_repor where rep_or = '" & rep & "'", gconDMIS
    rsREPOR.Open "select rep_or,niym,acct_no,invoice from CSMS_repor where rep_or = '" & txtRONO.Text & "'", gconDMIS
    If Not rsREPOR.EOF And Not rsREPOR.BOF Then
        If Null2String(rsREPOR!invoice) <> "" Then
            MsgBox "Warning: Repair Order is Already Released!" & vbCrLf & _
                 " Parts Issuance for this Repair Order must have a Reference Advanced Bill!", vbCritical, "Critical Issue!"
            If MsgBox("Would You Like to Continue?", vbQuestion + vbYesNo, "Continue...") = vbNo Then
                On Error Resume Next
                txtRONO.SetFocus
                Exit Sub
            Else
                MsgBox "Pls. Input Reference Number from Remarks Field..."
                On Error Resume Next
                txtRemarks.SetFocus
            End If
        End If
        txtCustName.Text = Null2String(rsREPOR!niym)
        txtCustCode.Text = Null2String(rsREPOR!ACCT_NO)
    Else
        txtCustName.Text = ""
        txtCustCode.Text = ""
    End If
End Sub

Sub InsertAdvanceBill()
    Dim ORDTRANDATE, ORDTRANNO, ORDTRANTYPE            As String
    Dim ORDITEMNO, ORDSTOCK_ORD, ORDSTOCK_SUP          As String
    Dim ORDTRANQTY                                     As Integer
    Dim ORDUNIT                                        As String
    Dim ORDTRANUCOST                                   As Double
    Dim ORDSTATUS, ORDIN_OUT                           As String
    Dim ORDTRANINVAMT                                  As Double

    Dim CurONHAND, CurSAFESTOCK, CurTISSQTY            As Integer
    Dim curRESSERVICE, curIssuances, PrevCurOrdQty     As Integer

    If txtTranType.Text = "RIV" Then
        Dim rsAdvanceBill                              As ADODB.Recordset
        Set rsAdvanceBill = New ADODB.Recordset
        rsAdvanceBill.Open "select PMIS_ORD_HD.rono,PMIS_ORD_HD.trandate,PMIS_ORD_HD.trantype,PMIS_ORD_HD.tranno,PMIS_TDAYTRAN.trantype,PMIS_TDAYTRAN.tranno,PMIS_TDAYTRAN.itemno,PMIS_TDAYTRAN.STOCK_ORD,PMIS_TDAYTRAN.tranqty,PMIS_TDAYTRAN.tranuprice from PMIS_Ord_Hd inner join PMIS_TDAYTRAN on PMIS_ORD_HD.TRANNO = PMIS_TDAYTRAN.TRANNO and PMIS_ORD_HD.TRANTYPE = PMIS_TDAYTRAN.TRANTYPE where PMIS_ORD_HD.[TYPE] = 'P' AND PMIS_ORD_HD.trantype = 'ADB' and PMIS_ord_hd.rono = '" & txtRONO.Text & "'", gconDMIS
        If Not rsAdvanceBill.EOF And Not rsAdvanceBill.BOF Then
            If MsgQuestionBox("Advance Bill for Repair Order: " & txtRONO.Text & " is Available " & vbCrLf & _
                              "Would you like to Insert this Transaction?", "Available Advance Bill") = True Then


                rsAdvanceBill.MoveFirst
                Do While Not rsAdvanceBill.EOF
                    ORDTRANDATE = N2Date2Null(txtTranDate.Text)
                    ORDTRANTYPE = "'RIV'"
                    ORDTRANNO = "'" & txtTranNo.Text & "'"
                    ORDITEMNO = N2Str2Null(Format(rsAdvanceBill!itemno, "0000"))
                    ORDSTOCK_ORD = N2Str2Null(rsAdvanceBill!STOCK_ORD)
                    ORDSTOCK_SUP = N2Str2Null(rsAdvanceBill!STOCK_ORD)
                    ORDTRANQTY = N2Str2IntZero(rsAdvanceBill!tranqty)
                    ORDTRANINVAMT = N2Str2Zero(rsAdvanceBill!TRANUPRICE)
                    ORDIN_OUT = "'O'"
                    ORDSTATUS = "'N'"

                    Set rsPartMas = New ADODB.Recordset
                    rsPartMas.Open "Select STOCKNO,onhand,sstock,resservice,TISSQTY,issuances from PMIS_PARTMAS where STOCKNO = " & ORDSTOCK_ORD & " AND ACTIVE = 'Y'", gconDMIS
                    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                        CurONHAND = N2Str2IntZero(rsPartMas!ONHAND)
                        CurSAFESTOCK = N2Str2IntZero(rsPartMas!SSTOCK)
                        CurTISSQTY = N2Str2IntZero(rsPartMas!TISSQTY)
                        curRESSERVICE = N2Str2IntZero(rsPartMas!RESSERVICE)
                        curIssuances = N2Str2IntZero(rsPartMas!issuances)

                        If CurONHAND <= 0 Then
                            Screen.MousePointer = 0
                            MsgSpeechBox "Part Number: " & Null2String(rsAdvanceBill!STOCK_ORD) & " is Out of Stock!"
                            If MsgQuestionBox("Warning: Error Has been encountered... Continue Anyway?", "Error Encountered") = False Then
                                Exit Sub
                            End If
                        End If
                        If ORDTRANQTY > CurONHAND Then
                            Screen.MousePointer = 0
                            MsgSpeechBox "Part Number " & Null2String(rsAdvanceBill!STOCKNO) & " Qty Ordered Exceeds Current Stock!" & vbCrLf & _
                                         "This Transaction will not be Included in RIV Transaction..."
                        Else
                            CurONHAND = CurONHAND - ORDTRANQTY
                        End If
                    Else
                        Screen.MousePointer = 0
                        MsgSpeechBox "Part Number: " & Null2String(rsAdvanceBill!STOCK_ORD) & " Not Found!"
                    End If

                    gconDMIS.Execute "insert into PMIS_TdayTran " & _
                                     "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice,lastupdate,usercode,status,in_out)" & _
                                   " values ('P'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                                   " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                                   " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                                   " " & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & ")"
                    cleargrid grdDetails
                    DoEvents
                    FillDetails
                    gconDMIS.Execute "update PMIS_Ord_Hd set" & _
                                   " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                                   " netinvamt = " & ORD_TOTINVAMT & _
                                   " where id = " & labID.Caption
                    gconDMIS.Execute "update PMIS_PARTMAS set" & _
                                   " onhand = " & CurONHAND & "," & _
                                   " TISSQTY = " & CurTISSQTY + ORDTRANQTY & ", " & _
                                   " issuances = " & curIssuances + ORDTRANQTY & _
                                   " where STOCKNO = " & ORDSTOCK_SUP
                    rsAdvanceBill.MoveNext
                Loop
            End If
        End If

        Set rsAdvanceBill = New ADODB.Recordset
        rsAdvanceBill.Open "select PMIS_ORD_HIST.rono,PMIS_ORD_HIST.trandate,PMIS_ORD_HIST.trantype,PMIS_ORD_HIST.tranno,PMIS_DAYTRAN.trantype,PMIS_DAYTRAN.tranno,PMIS_DAYTRAN.itemno,PMIS_DAYTRAN.STOCK_ORD,PMIS_DAYTRAN.tranqty,PMIS_DAYTRAN.tranuprice from PMIS_Ord_Hist inner join PMIS_DAYTRAN on PMIS_ORD_HIST.TRANNO = PMIS_DAYTRAN.TRANNO and PMIS_ORD_HIST.TRANTYPE = PMIS_DAYTRAN.TRANTYPE where PMIS_ORD_HIST.[TYPE] = 'P' AND PMIS_ORD_HIST.trantype = 'ADB' and PMIS_ord_hIST.rono = '" & txtRONO.Text & "'", gconDMIS
        If Not rsAdvanceBill.EOF And Not rsAdvanceBill.BOF Then
            If MsgQuestionBox("Advance Bill for Repair Order: " & txtRONO.Text & " is Available " & vbCrLf & _
                              "Would you like to Insert this Transaction?", "Available Advance Bill") = True Then

                rsAdvanceBill.MoveFirst
                Do While Not rsAdvanceBill.EOF
                    ORDTRANDATE = N2Date2Null(txtTranDate.Text)
                    ORDTRANTYPE = "'RIV'"
                    ORDTRANNO = "'" & txtTranNo.Text & "'"
                    ORDITEMNO = N2Str2Null(Format(rsAdvanceBill!itemno, "0000"))
                    ORDSTOCK_ORD = N2Str2Null(rsAdvanceBill!STOCK_ORD)
                    ORDSTOCK_SUP = N2Str2Null(rsAdvanceBill!STOCK_ORD)
                    ORDTRANQTY = N2Str2IntZero(rsAdvanceBill!tranqty)
                    ORDTRANINVAMT = N2Str2Zero(rsAdvanceBill!TRANUPRICE)
                    ORDIN_OUT = "'O'"
                    ORDSTATUS = "'N'"

                    Set rsPartMas = New ADODB.Recordset
                    rsPartMas.Open "Select STOCKNO,onhand,sstock,resservice,TISSQTY,issuances from PMIS_PARTMAS where STOCKNO = " & ORDSTOCK_ORD & " AND ACTIVE = 'Y'", gconDMIS
                    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                        CurONHAND = N2Str2IntZero(rsPartMas!ONHAND)
                        CurSAFESTOCK = N2Str2IntZero(rsPartMas!SSTOCK)
                        CurTISSQTY = N2Str2IntZero(rsPartMas!TISSQTY)
                        curRESSERVICE = N2Str2IntZero(rsPartMas!RESSERVICE)
                        curIssuances = N2Str2IntZero(rsPartMas!issuances)

                        If CurONHAND <= 0 Then
                            Screen.MousePointer = 0
                            MsgSpeechBox "Part Number: " & Null2String(rsAdvanceBill!STOCK_ORD) & " is Out of Stock!"
                            If MsgQuestionBox("Warning: Error Has been encountered... Continue Anyway?", "Error Encountered") = False Then
                                Exit Sub
                            End If
                        End If
                        If ORDTRANQTY > CurONHAND Then
                            Screen.MousePointer = 0
                            MsgSpeechBox "Part Number " & Null2String(rsAdvanceBill!STOCKNO) & " Qty Ordered Exceeds Current Stock!" & vbCrLf & _
                                         "This Transaction will not be Included in RIV Transaction..."
                        Else
                            CurONHAND = CurONHAND - ORDTRANQTY
                        End If
                    Else
                        Screen.MousePointer = 0
                        MsgSpeechBox "Part Number: " & Null2String(rsAdvanceBill!STOCK_ORD) & " Not Found!"
                    End If

                    gconDMIS.Execute "insert into PMIS_TdayTran " & _
                                     "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice,lastupdate,usercode,status,in_out)" & _
                                   " values ('P'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                                   " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                                   " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                                   " " & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & ")"
                    cleargrid grdDetails
                    DoEvents
                    FillDetails
                    gconDMIS.Execute "update PMIS_Ord_Hd set" & _
                                   " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                                   " netinvamt = " & ORD_TOTINVAMT & _
                                   " where id = " & labID.Caption
                    gconDMIS.Execute "update PMIS_PARTMAS set" & _
                                   " onhand = " & CurONHAND & "," & _
                                   " TISSQTY = " & CurTISSQTY + ORDTRANQTY & ", " & _
                                   " issuances = " & curIssuances + ORDTRANQTY & _
                                   " where STOCKNO = " & ORDSTOCK_SUP
                    rsAdvanceBill.MoveNext
                Loop
            End If
        End If
    End If
End Sub

Function SetSTOCKDESC(ppp As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select STOCKNO,STOCKDESC,srp,mac,dnp from PMIS_STOCKMAS where STOCKNO= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetSTOCKDESC = Null2String(rsPartMas!STOCKDESC)
        If txtTranType.Text = "DR" Then
            If cboChargeTo.Text = "PARTS CLAIM" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            End If
        Else
            If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                'If cboChargeTo.Text = "WARRANTY" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!DNP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                'ElseIf cboChargeTo.Text = "COMPANY" Then
                '   txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPARTMAS!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            End If
        End If
    Else
        If COUNTERTYPE = "ADB" Then
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "Select STOCKNUMBER,descriptio,srp from PMIS_DNPP where STOCKNUMBER= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                SetSTOCKDESC = Null2String(rsPartMas!DESCRIPTIO)
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            Else
                txtTranUPrice.Text = 0
                txtTranUCost.Text = 0
            End If
        Else
            txtTranUPrice.Text = 0
            txtTranUCost.Text = 0
        End If
    End If
End Function

Function SetSTOCKDESC2(pid As Variant)
    If COUNTERTYPE = "ADB" Then
        Set rsPartMas = New ADODB.Recordset
        'updating code:     JAA - 09202007
        rsPartMas.Open "Select PARTNUMBER,descriptio,dnpp,srp from PMIS_Dnpp where PARTNUMBER = " & N2Str2Null(cboTranPartNo.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then
            SetSTOCKDESC2 = Null2String(rsPartMas!DESCRIPTIO)
            txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
            'txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!DNPP))
        Else
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "Select id,STOCKDESC,srp,mac,dnp from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(cboTranPartNo.Text) & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                SetSTOCKDESC2 = Null2String(rsPartMas!STOCKDESC)
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            Else
                txtTranUPrice.Text = 0
                txtTranUCost.Text = 0
            End If
        End If
    Else
        If pid <> "" Then
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "Select id,STOCKDESC,srp,mac,dnp from PMIS_PARTMAS where id = " & pid & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                SetSTOCKDESC2 = Null2String(rsPartMas!STOCKDESC)
                If txtTranType.Text = "DR" Then
                    If cboChargeTo.Text = "PARTS CLAIM" Then
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                    Else
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                    End If
                Else
                    If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                        'If cboChargeTo.Text = "WARRANTY" Then
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!DNP))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                        'ElseIf cboChargeTo.Text = "COMPANY" Then
                        '   txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPARTMAS!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                    Else
                        Dim rsPRS_Header               As ADODB.Recordset
                        Dim rsPRS_Details              As ADODB.Recordset
                        Set rsPRS_Header = New ADODB.Recordset
                        Set rsPRS_Header = gconDMIS.Execute("Select * from PMIS_vw_PRS where REFPISNO = '" & cboRefPRSNo.Text & "'")
                        If Not rsPRS_Header.EOF And Not rsPRS_Header.BOF Then
                            Set rsPRS_Details = New ADODB.Recordset
                            Set rsPRS_Details = gconDMIS.Execute("Select * from PMIS_vw_PRS_Tran Where Tranno = " & N2Str2Null(rsPRS_Header!Tranno) & " AND STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text))
                            If Not rsPRS_Details.EOF And Not rsPRS_Details.BOF Then
                                txtTranQty.Text = N2Str2Zero(rsPRS_Details!tranqty)
                            Else
                            End If
                        Else
                        End If
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                    End If
                End If
            Else
                txtTranUPrice.Text = "0.00"
                txtTranUCost.Text = 0
            End If
        End If
    End If
End Function

Function SetSTOCKNO(pid As Variant)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select id,STOCKNO,srp,dnp,mac from PMIS_PARTMAS where id = " & pid & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetSTOCKNO = Null2String(rsPartMas!STOCKNO)
        If txtTranType.Text = "DR" Then
            If cboChargeTo.Text = "PARTS CLAIM" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            End If
        Else
            If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                'If cboChargeTo.Text = "WARRANTY" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!DNP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                'ElseIf cboChargeTo.Text = "COMPANY" Then
                '   txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPARTMAS!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            End If
        End If
    Else
        txtTranUPrice.Text = "0.00"
        txtTranUCost.Text = 0
    End If
End Function

Function SetPartIDSTOCKNO(DDD As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select id,STOCKNO from PMIS_PARTMAS where STOCKNO = '" & DDD & "' AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartIDSTOCKNO = Null2String(rsPartMas!ID)
        SetPartDetails DDD
    End If
End Function

Function SetPartIDDesc(DDD As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select id,STOCKDESC from PMIS_PARTMAS where ltrim(rtrim(STOCKDESC)) = '" & LTrim(RTrim(DDD)) & "' AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartIDDesc = Null2String(rsPartMas!ID)
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set rsPartMas = New ADODB.Recordset
        rsPartMas.Open "Select srp,STOCKNO,mac,dnp from PMIS_PARTMAS where STOCKNO = '" & ppp & "' AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then
            'SetPartPrice = Format(N2Str2Zero(rsPARTMAS!SRP), MAXIMUM_DIGIT)
            If txtTranType.Text = "DR" Then
                If cboChargeTo.Text = "PARTS CLAIM" Then
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                Else
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                End If
            Else
                If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                    'If cboChargeTo.Text = "WARRANTY" Then
                    'SetPartPrice = ToDoubleNumber(N2Str2Zero(rsPartMas!DNP))
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac) * 1.12)
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                    'ElseIf cboChargeTo.Text = "COMPANY" Then
                    '   SetPartPrice = ToDoubleNumber(N2Str2Zero(rsPARTMAS!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                Else
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                End If
            End If
        End If
        SetPartDetails ppp
    End If
End Function

Sub SetPartDetails(XXX As String)
    Dim rsPartMas                                      As ADODB.Recordset
    Set rsPartMas = New ADODB.Recordset
    '==========================================================================================================
    'Updating Code:     JAA - 10252007
    'Set rsPartMas = gconDMIS.Execute("Select * from PMIS_PartMas where PartNo = '" & XXX & "' AND ACTIVE = 'Y'")
    Set rsPartMas = gconDMIS.Execute("Select * from PMIS_StockMas where TYPE = 'P' and StockNo = '" & XXX & "' AND ACTIVE = 'Y'")
    '==========================================================================================================
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        If N2Str2Zero(rsPartMas!ONHAND) > 0 Then chkAvailableOnStock.Value = 1 Else chkAvailableOnStock.Value = 0
        optLocalPurchase.Value = False: optImported.Value = False: optConsigned.Value = False
        optGenuine.Value = False: optNonGenuine.Value = False
        If Null2String(rsPartMas!PartsOrigin) = "M" Then
            optImported.Value = True
        End If
        If Null2String(rsPartMas!PartsOrigin) = "L" Then
            optLocalPurchase.Value = True
        End If
        If Null2String(rsPartMas!PartsOrigin) = "C" Then
            optConsigned.Value = True
        End If
        If Null2String(rsPartMas!Genuine) = "Y" Then
            optGenuine.Value = True
        Else
            optNonGenuine.Value = True
        End If
        txtModelCode.Text = Null2String(rsPartMas!modelcode)
    Else
        optLocalPurchase.Value = False
        optImported.Value = False
        optConsigned.Value = False
        optGenuine.Value = False
        optNonGenuine.Value = False
        txtModelCode.Text = ""
    End If
End Sub

Sub InitParts()
    txtTranItemNo.Text = Format(kcnt + 1, "0000")
    cboTranPartNo.Text = ""
    txtTranDescription.Text = ""
    txtTranQty.Text = 1
    txtTranUCost.Text = "0.00"
    txtTranUPrice.Text = "0.00"
    txtTranTotalAmt.Text = "0.00"
    If COUNTERTYPE = "ADB" Then
        labTranUCost.Visible = True: txtTranUCost.Visible = True
    Else
        labTranUCost.Visible = False: txtTranUCost.Visible = False
    End If
    Check1.Enabled = False
End Sub

Function StorePartsEntry(ByVal ID As Variant)
    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select id,STOCK_ORD,STOCK_SUP,tranqty,itemno,tranuprice,tranucost from PMIS_TdayTran where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        labDetID.Caption = rsTdayTran!ID
        labPartNo.Caption = Null2String(rsTdayTran!STOCK_ORD)
        labPrevOrdQty.Caption = N2Str2IntZero(rsTdayTran!tranqty)
        txtTranItemNo.Text = Null2String(rsTdayTran!itemno)
        cboTranPartNo.Text = Null2String(rsTdayTran!STOCK_ORD)
        txtTranDescription.Text = SetSTOCKDESC(rsTdayTran!STOCK_ORD)
        txtTranQty.Text = N2Str2IntZero(rsTdayTran!tranqty)
        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsTdayTran!TRANUPRICE))
        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsTdayTran!TRANUCOST))
        txtTranTotalAmt.Text = ToDoubleNumber(N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE))
    End If
    If COUNTERTYPE = "ADB" Then
        labTranUCost.Visible = True: txtTranUCost.Visible = True
    Else
        labTranUCost.Visible = False: txtTranUCost.Visible = False
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    PMIS_ORDER_SHOW = False: Set frmPMISCustomerOrder_CSMS = Nothing
    UnloadForm Me
End Sub

Private Sub grdDetails_DblClick()
    Dim fild                                           As String
    If Null2String(rsOrd_Hd!Status) = "C" Then
        MsgSpeech "Transactions are Already Cancelled and cannot be Change"
        MsgBoxXP "Transactions are Already Cancelled" & vbCrLf & _
                 "and cannot be Change", "Edit Not Allowed!", XP_OKOnly, msg_Information
    ElseIf Null2String(rsOrd_Hd!Status) = "B" Then
        MsgSpeech "Transactions are Already Billed-Out and cannot be Change"
        MsgBoxXP "Transactions are Already Billed-Out" & vbCrLf & _
                 "and cannot be Change", "Edit Not Allowed!", XP_OKOnly, msg_Information
    ElseIf Null2String(rsOrd_Hd!Status) = "P" Then
        MsgSpeech "Transactions are Already Posted and cannot be Change"
        MsgBoxXP "Transactions are Already Posted" & vbCrLf & _
                 "and cannot be Change", "Edit Not Allowed!", XP_OKOnly, msg_Information
    Else
        grdDetails.Row = grdDetails.Row
        grdDetails.Col = 0
        fild = grdDetails.Text
        If fild <> "" And fild <> "No Entry" Then
            AddorEdit = "EDIT"
            cmdTranDelete.Enabled = True
            BringToFront
            StorePartsEntry (fild)
        Else
            MsgSpeechBox "No Entry on Parts!"
            Exit Sub
        End If
    End If
End Sub

Sub SendToBack()
    cmdAddTran.ZOrder 1
    cmdAddTran.Visible = False
    fraAddTran.ZOrder 1
    fraAddTran.Visible = False
    fraAddTran.Enabled = False
    cmdSignatories.ZOrder 1
    cmdSignatories.Visible = False
    fraSignatories.ZOrder 1
    fraSignatories.Visible = False
End Sub

Sub BringToFront()
    cmdAddTran.ZOrder 0
    cmdAddTran.Visible = True
    fraAddTran.ZOrder 0
    fraAddTran.Visible = True
    fraAddTran.Enabled = True
End Sub

Private Sub optCASH_Click()
    COUNTERTYPE = "CSH"
End Sub

Private Sub optCHARGE_Click()
    COUNTERTYPE = "CHG"
End Sub

Private Sub Timer1_Timer()
    If labPosted.Caption <> "" Then
        If labPosted.Visible = True Then
            labPosted.Visible = False
        Else
            labPosted.Visible = True
        End If
    End If
End Sub

Sub SetCustomer()
    Dim rsCUSTOMER                                     As ADODB.Recordset
    Set rsCUSTOMER = New ADODB.Recordset
    Set rsCUSTOMER = gconDMIS.Execute("Select * from ALL_Customer where CusCde = '" & txtCustCode.Text & "'")
    If Not rsCUSTOMER.EOF And Not rsCUSTOMER.BOF Then
        txtCustName.Text = Null2String(rsCUSTOMER!acctname) & vbCrLf & Null2String(rsCUSTOMER!customeradd) & vbCrLf & Null2String(rsCUSTOMER!City)
    End If
End Sub

Private Sub txtDS_Desc1_Change()
    If Len(txtDS_Desc1.Text) = 1 Then
        If txtDS_Desc1.Text = "D" Then
            txtDS_Desc1.Text = "DISCOUNT"
        End If
    End If
End Sub

Private Sub txtDS1_Change()
    If NumericVal(txtDS1.Text) <> 0 Then
        If txtDS_Desc1.Text = "" Then
            txtDS_Desc1.Text = "DISCOUNT"
        End If
        txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
        txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
    Else
        txtDS_Desc1.Text = ""
        txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
        txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
    End If
End Sub

Private Sub txtDS1_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtDS1_LostFocus()
    If NumericVal(txtDS1.Text) <> 0 Then
        txtDS_Desc1.Text = "DISCOUNT"
        txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
        txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
    Else
        txtDS_Desc1.Text = ""
        txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
        txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
    End If
End Sub

Private Sub txtReferencePIS_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtRemarks_GotFocus()
    If txtRemarks.Text = "Pls Type Your Message Here!" Then txtRemarks.Text = ""
End Sub

Private Sub txtRONO_LostFocus()
    Dim RONOStr                    As String
    RONOStr = txtRONO.Text
    If Left(RONOStr, 2) = "R-" Then
       RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
    Else
       RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
    End If
    txtRONO.Text = RONOStr
    SetCustInfo (RONOStr)
End Sub

Private Sub txtTranDate_LostFocus()
    txtTranDate.Text = Format(txtTranDate.Text, "SHORT DATE")
End Sub

Private Sub txtTranNo_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txttranQty_Change()
    If txtTranQty.Text <> "" Then
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
    End If
End Sub

Private Sub txtTranQty_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtTranQty_LostFocus()
    If txtTranQty.Text <> "" Then
        txtTranTotalAmt.Text = Format(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text), MAXIMUM_DIGIT)
    Else
        txtTranQty.Text = 1
        txtTranTotalAmt.Text = Format(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text), MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtTranQty_GotFocus()
    If NumericVal(txtTranQty.Text) = 1 Then txtTranQty.Text = ""
End Sub

Private Sub txtTranTotalAmt_Change()
    txtTranTotalAmt.Text = Format(txtTranTotalAmt.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtTranUPrice_Change()
    If txtTranUPrice.Text <> "" Then
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
    End If
End Sub

Private Sub txtUnitPrice_LostFocus()
    If txtTranUPrice.Text = "" Then txtTranUPrice.Text = 0
    txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
End Sub

Private Sub grdDetails_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtTranUPrice_GotFocus()
    If NumericVal(txtTranUPrice.Text) = 0 Then txtTranUPrice.Text = ""
End Sub

Private Sub txtTranUPrice_KeyPress(KeyCode As Integer)
    If (KeyCode < 48 Or KeyCode > 57) And KeyCode <> 110 And KeyCode <> 46 Then
        KeyCode = 0
    End If
End Sub

Private Sub txtTranUPrice_LostFocus()
    txtTranUPrice.Text = Format(txtTranUPrice.Text, MAXIMUM_DIGIT)
End Sub

'SEARCH MODULE
Private Sub lstOrd_Hd_GotFocus()
    rsOrd_Hd.Bookmark = rsFind(rsOrd_Hd.Clone, "tranno", lstOrd_Hd.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstOrd_Hd_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If optTranno.Value = True Then
        rsOrd_Hd.Bookmark = rsFind(rsOrd_Hd.Clone, "tranno", lstOrd_Hd.SelectedItem.SubItems(1)).Bookmark
    Else
        rsOrd_Hd.Bookmark = rsFind(rsOrd_Hd.Clone, "tranno", lstOrd_Hd.SelectedItem.SubItems(1)).Bookmark
    End If
    StoreMemVars
End Sub

Private Sub lstOrd_Hd_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstOrd_Hd
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstOrd_Hd_DblClick()
    If cmdEdit.Enabled = True Then cmdEdit.Value = True
End Sub

Private Sub lstOrd_Hd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If optTranno.Value = True Then
        If Trim(textSearch.Text) = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    Else
        Dim RONOStr                    As String
        RONOStr = textSearch.Text
        If Left(RONOStr, 2) = "R-" Then
           RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
        Else
           RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
        End If
        'txtRONO.Text = RONOStr
        If Trim(textSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (RONOStr)
    End If
End Sub

Sub FillGrid()
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set rsOrd_Hd = New ADODB.Recordset
    Set rsOrd_Hd = gconDMIS.Execute("select Tranno,tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & COUNTERTYPE & "' order by Tranno asc")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True: Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd: lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False
    Set rsOrd_Hd = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsOrd_Hd = gconDMIS.Execute("select tranno, tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & COUNTERTYPE & "' and tranno like '" & XXX & "%'")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True: Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd: lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set rsOrd_Hd = New ADODB.Recordset
    Set rsOrd_Hd = gconDMIS.Execute("select rono,tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & COUNTERTYPE & "' and rono is not null order by tranno asc")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True: Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd: lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set rsOrd_Hd = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsOrd_Hd = gconDMIS.Execute("select Rono, tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & COUNTERTYPE & "' and rono like '" & XXX & "%' order by tranno asc")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True: Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd: lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstOrd_Hd.ListItems.Count > 0 And lstOrd_Hd.Enabled = True Then: lstOrd_Hd.SetFocus
    End If
End Sub

Private Sub optRONo_Click()
    lstOrd_Hd.ColumnHeaders(1).Text = "RO Number"
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optTranno_Click()
    lstOrd_Hd.ColumnHeaders(1).Text = "Tran. No."
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub
