VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmPMISTrans_ADB_Issuances_Acc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issuances Against Advance Bill"
   ClientHeight    =   7155
   ClientLeft      =   1110
   ClientTop       =   2520
   ClientWidth     =   11505
   ForeColor       =   &H00DEDFDE&
   Icon            =   "frmPMISTrans_ADB_Issuances_Acc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11505
   Begin VB.PictureBox fraAddTran 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3555
      Left            =   3450
      ScaleHeight     =   3525
      ScaleWidth      =   6735
      TabIndex        =   42
      Top             =   1470
      Width           =   6765
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
         Height          =   270
         Left            =   3900
         TabIndex        =   138
         Top             =   30
         Width           =   2595
      End
      Begin VB.CommandButton Command1 
         Caption         =   "View All Stocks For This RO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         TabIndex        =   128
         Top             =   2310
         Width           =   2535
      End
      Begin VB.TextBox TXT_ADB_FILL 
         Alignment       =   2  'Center
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
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   126
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   750
         Width           =   855
      End
      Begin VB.TextBox TXT_AVL_ISSUE 
         Alignment       =   2  'Center
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
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   124
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   1860
         Width           =   855
      End
      Begin VB.TextBox TXT_CURRONHAND 
         Alignment       =   2  'Center
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
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   121
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   1170
         Width           =   855
      End
      Begin VB.TextBox TXT_ADB_ISSUE 
         Alignment       =   2  'Center
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
         Left            =   4800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   119
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   330
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   3150
         TabIndex        =   105
         Top             =   1890
         Width           =   285
      End
      Begin VB.CommandButton cmdTranDelete 
         Caption         =   "&Delete"
         CausesValidation=   0   'False
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Delete Entry"
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton cmdTranCancel 
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":0D47
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":0E99
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Cancel Entry"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtTranDescription 
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         Top             =   1110
         Width           =   3675
      End
      Begin VB.TextBox txtTranTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         Top             =   2190
         Width           =   1665
      End
      Begin VB.TextBox txtTranUPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         Top             =   1860
         Width           =   1665
      End
      Begin VB.TextBox txtTranQty 
         Alignment       =   2  'Center
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
         Top             =   1500
         Width           =   705
      End
      Begin VB.TextBox txtTranItemNo 
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   16
         ToolTipText     =   "Type item number (e.g. 0001)"
         Top             =   60
         Width           =   765
      End
      Begin VB.ComboBox cboTranPartNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         ItemData        =   "frmPMISTrans_ADB_Issuances_Acc.frx":11D7
         Left            =   1440
         List            =   "frmPMISTrans_ADB_Issuances_Acc.frx":11D9
         Sorted          =   -1  'True
         TabIndex        =   17
         Text            =   "Combo1"
         ToolTipText     =   "Select Part Number from the list."
         Top             =   480
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
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   480
         Width           =   585
      End
      Begin VB.Frame fraCostToCost 
         Height          =   405
         Left            =   2190
         TabIndex        =   106
         Top             =   1410
         Width           =   1575
         Begin VB.CheckBox Check1 
            Caption         =   "Cost to Cost"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   107
            Top             =   150
            Width           =   1395
         End
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
         TabIndex        =   103
         Text            =   "1000.00"
         ToolTipText     =   "Input price of item. Do not use comma and peso sign (e.g.300, 26)"
         Top             =   1500
         Width           =   945
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":11DB
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":132D
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Save Entry"
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "On hand"
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
         Index           =   2
         Left            =   3960
         TabIndex        =   125
         Top             =   1230
         Width           =   810
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Filled Qty"
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
         Index           =   1
         Left            =   3900
         TabIndex        =   123
         Top             =   810
         Width           =   870
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Available For Issuance"
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
         Index           =   0
         Left            =   3990
         TabIndex        =   122
         Top             =   1590
         Width           =   2505
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ADB Qty"
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
         Left            =   3960
         TabIndex        =   120
         Top             =   420
         Width           =   765
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
         TabIndex        =   104
         Top             =   1530
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
         TabIndex        =   57
         Top             =   1860
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
         TabIndex        =   55
         Top             =   1890
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
         TabIndex        =   49
         Top             =   2250
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
         TabIndex        =   48
         Top             =   1890
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
         TabIndex        =   47
         Top             =   1530
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
         TabIndex        =   46
         Top             =   510
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
         TabIndex        =   45
         Top             =   150
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
         TabIndex        =   44
         Top             =   870
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
         TabIndex        =   56
         Top             =   1860
         Width           =   975
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "View All Stocks For RO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   135
      Top             =   6660
      Width           =   2565
   End
   Begin VB.PictureBox Picture6 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   109
      Top             =   6810
      Width           =   11505
      Begin VB.Label LAB_ADB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   8400
         TabIndex        =   113
         Top             =   0
         Width           =   3075
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
      TabIndex        =   73
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
         Top             =   30
         Width           =   1455
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   6645
      Left            =   60
      TabIndex        =   66
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
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   1260
         Width           =   2475
      End
      Begin VB.OptionButton optRONo 
         Caption         =   "RO Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   68
         Top             =   630
         Width           =   2385
      End
      Begin VB.OptionButton optTranno 
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   67
         Top             =   390
         Value           =   -1  'True
         Width           =   2385
      End
      Begin MSComctlLib.ListView lstOrd_Hd 
         Height          =   4965
         Left            =   30
         TabIndex        =   70
         Top             =   1620
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   8758
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":167D
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
      Begin VB.OptionButton Option1 
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   110
         Top             =   900
         Width           =   2205
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
         TabIndex        =   71
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
      TabIndex        =   41
      Top             =   3150
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
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   2640
      ScaleHeight     =   915
      ScaleWidth      =   8835
      TabIndex        =   87
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":17DF
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":1931
         Style           =   1  'Graphical
         TabIndex        =   90
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":1C97
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":1DE9
         Style           =   1  'Graphical
         TabIndex        =   91
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":214F
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":22A1
         Style           =   1  'Graphical
         TabIndex        =   97
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":25DB
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":272D
         Style           =   1  'Graphical
         TabIndex        =   98
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":2A52
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":2BA4
         Style           =   1  'Graphical
         TabIndex        =   92
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":2F00
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":3052
         Style           =   1  'Graphical
         TabIndex        =   93
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":3365
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":34B7
         Style           =   1  'Graphical
         TabIndex        =   89
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":3807
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":3959
         Style           =   1  'Graphical
         TabIndex        =   88
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":3CB7
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":3E09
         Style           =   1  'Graphical
         TabIndex        =   94
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":4103
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":4255
         Style           =   1  'Graphical
         TabIndex        =   95
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":45AD
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":46FF
         Style           =   1  'Graphical
         TabIndex        =   96
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
      TabIndex        =   84
      Top             =   5895
      Width           =   2220
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":4A5E
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":4BB0
         Style           =   1  'Graphical
         TabIndex        =   85
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":4EEE
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":5040
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   795
      End
   End
   Begin VB.PictureBox fraSignatories 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   4185
      ScaleHeight     =   2325
      ScaleWidth      =   4380
      TabIndex        =   50
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
         Left            =   3000
         MouseIcon       =   "frmPMISTrans_ADB_Issuances_Acc.frx":5390
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":54E2
         Style           =   1  'Graphical
         TabIndex        =   83
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
         Top             =   90
         Width           =   1395
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
      Top             =   0
      Width           =   8715
      Begin VB.CommandButton CMD_ADD_RO 
         Caption         =   "..."
         Height          =   345
         Left            =   2700
         TabIndex        =   114
         Top             =   870
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   ".."
         Height          =   345
         Left            =   2700
         TabIndex        =   111
         Top             =   30
         Width           =   375
      End
      Begin VB.CommandButton cmdEditTranDate 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   2700
         TabIndex        =   108
         Top             =   450
         Width           =   375
      End
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
         TabIndex        =   100
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
            TabIndex        =   102
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
            TabIndex        =   101
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
      Begin VB.CommandButton cmdPISNum 
         Caption         =   "..."
         Height          =   375
         Left            =   8310
         TabIndex        =   79
         Top             =   60
         Width           =   255
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
         Top             =   450
         Width           =   1485
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
         Picture         =   "frmPMISTrans_ADB_Issuances_Acc.frx":5848
         ScaleHeight     =   405
         ScaleWidth      =   435
         TabIndex        =   59
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
            TabIndex        =   60
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
         Left            =   3720
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
         Left            =   3870
         MaxLength       =   7
         TabIndex        =   4
         ToolTipText     =   "Type the transaction terms."
         Top             =   450
         Width           =   1005
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
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   0
         ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
         Top             =   30
         Width           =   1485
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
         TabIndex        =   58
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
            TabIndex        =   63
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
            TabIndex        =   62
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
            TabIndex        =   61
            Top             =   60
            Width           =   1395
         End
      End
      Begin VB.TextBox txtRONO 
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
         Top             =   840
         Width           =   1485
      End
      Begin VB.TextBox txtPRtranno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   30
         Visible         =   0   'False
         Width           =   1485
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
         Left            =   5640
         TabIndex        =   1
         Text            =   "PIWGC06H360"
         ToolTipText     =   "Type Reference PIS Number"
         Top             =   60
         Width           =   1785
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
         TabIndex        =   99
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
         Left            =   4770
         TabIndex        =   72
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
         Left            =   2370
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         Left            =   3150
         TabIndex        =   37
         Top             =   900
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
         Left            =   3210
         TabIndex        =   36
         Top             =   510
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
         TabIndex        =   35
         Top             =   480
         Width           =   1005
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
         TabIndex        =   34
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
         Left            =   120
         TabIndex        =   33
         Top             =   60
         Width           =   1065
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
         Top             =   900
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
         Left            =   7200
         TabIndex        =   29
         Top             =   90
         Width           =   1425
      End
   End
   Begin VB.PictureBox pic_viewStockADB 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   2100
      ScaleHeight     =   4665
      ScaleWidth      =   8925
      TabIndex        =   129
      Top             =   1260
      Visible         =   0   'False
      Width           =   8955
      Begin VB.OptionButton otp_RIVADB 
         Caption         =   "Show Issuances To ADB "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   137
         Top             =   480
         Width           =   2265
      End
      Begin VB.OptionButton otp_ADB 
         Caption         =   "Show ADB Issuances"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   136
         Top             =   480
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.TextBox TXT_ROVIEW 
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
         Left            =   5880
         TabIndex        =   133
         Top             =   30
         Width           =   1965
      End
      Begin VB.CommandButton Command4 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8520
         TabIndex        =   130
         Top             =   30
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3795
         Left            =   30
         TabIndex        =   131
         Top             =   810
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6694
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tran#"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Part Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Part Name"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Onhand"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "ADB Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Issused To ADB"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Command5 
         Caption         =   "View"
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
         Left            =   7860
         TabIndex        =   134
         Top             =   30
         Width           =   675
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   405
         Left            =   0
         TabIndex        =   132
         Top             =   0
         Width           =   8925
         _Version        =   655364
         _ExtentX        =   15743
         _ExtentY        =   714
         _StockProps     =   14
         Caption         =   "Showing All ADB Stock For this RO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox PIC_ADD_RO 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5025
      Left            =   3570
      ScaleHeight     =   4995
      ScaleWidth      =   6435
      TabIndex        =   115
      Top             =   570
      Visible         =   0   'False
      Width           =   6465
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   118
         Top             =   480
         Width           =   3525
      End
      Begin VB.CommandButton cmdClose_1 
         Caption         =   "..."
         Height          =   345
         Left            =   5940
         TabIndex        =   117
         Top             =   510
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4035
         Left            =   30
         TabIndex        =   116
         Top             =   900
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   7117
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ADB#"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RO#"
            Object.Width           =   2540
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   405
         Left            =   0
         TabIndex        =   127
         Top             =   0
         Width           =   6435
         _Version        =   655364
         _ExtentX        =   11351
         _ExtentY        =   714
         _StockProps     =   14
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
   End
End
Attribute VB_Name = "frmPMISTrans_ADB_Issuances_Acc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSORD_HD                                           As ADODB.Recordset
Dim RSTDAYTRAN                                         As ADODB.Recordset
Dim RSPARTMAS                                          As ADODB.Recordset
Dim RSSALESMAN                                         As ADODB.Recordset
Dim RSCUNTER                                           As ADODB.Recordset
Dim RSPROFILE                                          As ADODB.Recordset
Dim RSREPOR                                            As ADODB.Recordset
Dim RSCUSTOMER                                         As ADODB.Recordset

Dim KCNT                                               As Integer
Dim ADDOREDIT                                          As String
Dim ORD_TOTUPRICE                                      As Double
Dim ORD_TOTINVAMT                                      As Double
Dim ORD_TOTVAT                                         As Double
Dim ORD_TOTQTY                                         As Double
Dim PREVORDTYPE                                        As String
Dim PREVORDNO                                          As String
Dim REPOR_STATUS                                       As String
Dim LOCALACESS                                         As String
Private WithEvents FRM_SERIES                          As frmPMISTrans_ADB_Issuances_PISFormation
Attribute FRM_SERIES.VB_VarHelpID = -1


Sub BringToFront()
    Picture1.Enabled = False
    fraDetails.Enabled = False
    fraAddTran.ZOrder 0
    fraAddTran.Visible = True
    fraAddTran.Enabled = True
End Sub


Private Sub cboRefPRSNo_Click()
    cboRefPRSNo_LostFocus
End Sub

Private Sub cboRefPRSNo_GotFocus()
    Dim rsPRS                                          As ADODB.Recordset
    Dim rsPRS_HDDup                                    As ADODB.Recordset
    Set rsPRS = New ADODB.Recordset
    rsPRS.Open "SELECT TRANNO,REFPISNO FROM PMIS_VW_PRS WHERE TYPE = 'A' AND SALES_ORIGIN = 'S' ORDER BY TRANNO ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPRS.EOF And Not rsPRS.BOF Then
        rsPRS.MoveFirst: cboRefPRSNo.Clear
        Do While Not rsPRS.EOF
            Set rsPRS_HDDup = New ADODB.Recordset
            rsPRS_HDDup.Open "SELECT REFPISNO FROM PMIS_ORD_HD WHERE TRANTYPE <> 'ARS' AND [TYPE] = 'A' AND REFPRSNO = '" & Null2String(rsPRS!REFPISNO) & "'", gconDMIS
            If Not rsPRS_HDDup.EOF And Not rsPRS_HDDup.BOF Then
            Else
                cboRefPRSNo.AddItem Null2String(rsPRS!REFPISNO)
            End If
            rsPRS.MoveNext
        Loop
    End If
End Sub

Private Sub cboRefPRSNo_LostFocus()
    If ADDOREDIT = "ADD" Then
        Dim rsRR_HDDup                                 As ADODB.Recordset
        Set rsRR_HDDup = New ADODB.Recordset
        rsRR_HDDup.Open "select refpisno,tranno from PMIS_Ord_Hd where [TYPE] = 'A' AND refprsno = '" & cboRefPRSNo.Text & "'", gconDMIS
        If Not rsRR_HDDup.EOF And Not rsRR_HDDup.BOF Then
            MsgBox "ARS Number Already Received", vbInformation, "Invalid ARS Number"
            Exit Sub
        Else
            Set rsRR_HDDup = New ADODB.Recordset
            rsRR_HDDup.Open "select tranno,DS1,custname,custcode,rono from PMIS_vw_PRS where [TYPE] = 'A' AND refpisno = '" & cboRefPRSNo.Text & "'", gconDMIS
            If (rsRR_HDDup.EOF Or rsRR_HDDup.BOF) Then
                MsgSpeechBox "Invalid Accessories Requisition Number!"
            End If
        End If
    End If
End Sub

Private Sub cboSMName_Click()
    Set RSSALESMAN = New ADODB.Recordset
    RSSALESMAN.Open "select empno,signname from PMIS_vw_SalesMan where signname = " & N2Str2Null(cboSMName.Text), gconDMIS
    If Not RSSALESMAN.EOF And Not RSSALESMAN.BOF Then
        cboSalesMan.Text = Null2String(RSSALESMAN!empno)
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

Private Sub cboTranPartNo_GotFocus()
    VBComBoBoxDroppedDown cboTranPartNo
End Sub

Private Sub cboTranPartNo_LostFocus()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
    End If
End Sub

Private Sub Check1_Click()
    If Module_Access(LOGID, "APPLY PARTS COST TO COST AMOUNT", "SYSTEM") = False Then Check1.Value = 0: Exit Sub
    If Check1.Value = 1 Then
        txtTranUPrice.Text = txtTranUCost.Text
    Else
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
    End If
End Sub

Function CheckIfROBilled(xxx As String) As String
    Dim rsRo_det                                       As ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("SELECT INVOICE FROM CSMS_REPOR WHERE INVOICE IS NOT NULL AND REP_OR = " & N2Str2Null(xxx))
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        CheckIfROBilled = UCase(Null2String(rsRo_det!Invoice))
    End If
    Set rsRo_det = Nothing
End Function


Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", LOCALACESS) = False Then Exit Sub
    ADDOREDIT = "ADD"
    InitMemVars
    PisValidation
    Command3.Enabled = True
End Sub

Private Sub cmdAddTran_Click()
    SendToBack
    fraAddTran.Visible = True
    fraAddTran.ZOrder 0
    fraAddTran.Enabled = True
    ADDOREDIT = "ADD"
    cmdTranDelete.Enabled = False
    InitCbo
    InitParts
    On Error Resume Next
    cboTranPartNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    fraDetails.Enabled = True
    txtTranDate.Enabled = False
    StoreMemvars
    txtPRtranno.Visible = False
    Command3.Enabled = False
    cmdClose_1_Click
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", LOCALACESS) = False Then Exit Sub

    On Error GoTo Errorcode:

    If LOGLEVEL <> "ADM" Then
        MsgBox "Warning: Your account is not allowed to cancel this transaction!", vbCritical, "Error"
        Exit Sub
    End If

    If MsgQuestionBox("Are you sure you want to Cancel this Transaction?", "Cancel Transaction") = True Then
        Dim PCURONHAND                                 As Long
        Dim PCurTISSQTY                                As Long
        Dim PCURISSUANCES                              As Long
        Dim rsTdaytranDup                              As ADODB.Recordset
        Dim rsPartmasDup                               As ADODB.Recordset

        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "SELECT ID,TRANTYPE,TRANNO,STOCK_ORD,TRANQTY FROM PMIS_TDAYTRAN WHERE [TYPE] = 'A' AND TRANNO = " & N2Str2Null(RSORD_HD!TRANNO) & " AND TRANTYPE = " & N2Str2Null(RSORD_HD!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
            rsTdaytranDup.MoveFirst
            Do While Not rsTdaytranDup.EOF
                Set rsPartmasDup = New ADODB.Recordset
                rsPartmasDup.Open "SELECT STOCKNO,ONHAND,TISSQTY,TISSQTY,ISSUANCES,REQSERVED,S_REQSERVED FROM PMIS_STOCKMAS WHERE TYPE='A' AND STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD) & " AND ACTIVE = 'Y'", gconDMIS
                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                    PCURONHAND = N2Str2IntZero(rsPartmasDup!ONHAND) + N2Str2Zero(rsTdaytranDup!TRANQTY)
                    PCurTISSQTY = N2Str2IntZero(rsPartmasDup!TISSQTY) - N2Str2Zero(rsTdaytranDup!TRANQTY)
                    PCURISSUANCES = N2Str2IntZero(rsPartmasDup!ISSUANCES) - N2Str2Zero(rsTdaytranDup!TRANQTY)
                    If Null2String(RSORD_HD!STATUS) = "P" Then
                        If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                            SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                          " REQSERVED = " & N2Str2IntZero(rsPartmasDup!REQSERVED) - N2Str2Zero(rsTdaytranDup!TRANQTY) & _
                                          " WHERE TYPE='M' AND STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT
                            'NEW LOG AUDIT-------------------------------------------------
                            Call NEW_LogAudit("E", "ACCESSORIES MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "ACCESSORIES NO: " & Null2String(N2Str2Null(rsTdaytranDup!STOCK_ORD)), COUNTERTYPE, "")
                            'NEW LOG AUDIT-------------------------------------------------
                        Else
                            SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                          " S_REQSERVED = " & N2Str2IntZero(rsPartmasDup!S_REQSERVED) - N2Str2Zero(rsTdaytranDup!TRANQTY) & _
                                          " WHERE TYPE='M' AND STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT
                            'NEW LOG AUDIT-------------------------------------------------
                            Call NEW_LogAudit("E", "ACCESSORIES MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "ACCESSORIES NO: " & Null2String(N2Str2Null(rsTdaytranDup!STOCK_ORD)), COUNTERTYPE, "")
                            'NEW LOG AUDIT-------------------------------------------------
                        End If
                        SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                      " ONHAND = " & PCURONHAND & "," & _
                                      " TISSQTY = " & PCurTISSQTY & "," & _
                                      " ISSUANCES = " & PCURISSUANCES & "," & _
                                      " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                                      " LASTUPDATE = '" & LOGDATE & "'" & _
                                      " WHERE TYPE='M' AND STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                        gconDMIS.Execute SQL_STATEMENT
                        'NEW LOG AUDIT-------------------------------------------------
                        Call NEW_LogAudit("E", "ACCESSORIES MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "ACCESSORIES NO: " & Null2String(N2Str2Null(rsTdaytranDup!STOCK_ORD)), "", "")
                        'NEW LOG AUDIT-------------------------------------------------
                    End If
                    SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET" & _
                                  " STATUS = 'C'," & _
                                  " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                                  " LASTUPDATE = '" & LOGDATE & "'" & _
                                  " WHERE ID = " & rsTdaytranDup!ID
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "C", LOCALACESS, SQL_STATEMENT, labID, "ACCESSORIES", txtTranNo, COUNTERTYPE, ""
                End If

                rsTdaytranDup.MoveNext
            Loop
        End If
        SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                      " STATUS = 'C'," & _
                      " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                      " LASTUPDATE = '" & LOGDATE & "'" & _
                      " WHERE ID = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "C", LOCALACESS, SQL_STATEMENT, labID, "ACCESSORIES", txtTranNo, COUNTERTYPE, ""
        rsRefresh
        On Error Resume Next
        RSORD_HD.Find "id =" & labID.Caption
        StoreMemvars
    End If
    Set rsTdaytranDup = Nothing
    Set rsPartmasDup = Nothing
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdClose_1_Click()
    PIC_ADD_RO.Visible = False
    Frame1.Enabled = True
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", LOCALACESS) = False Then Exit Sub
    ADDOREDIT = "EDIT"
    PREVORDTYPE = txtTranType.Text
    PREVORDNO = Format(txtTranNo.Text, "000000")
    grdDetails.Enabled = False
    cmdEditTranDate.Enabled = True
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    txtTranDate.Enabled = False
    On Error Resume Next
    txtCustName.SetFocus
    Command3.Enabled = True
End Sub

Private Sub cmdEditTranDate_Click()
    If Function_Access(LOGID, "Acess_SYSTEM", LOCALACESS) = False Then Exit Sub
    txtTranDate.Enabled = True
    txtTranDate.Locked = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    MsgBox COUNTERTYPE
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    RSORD_HD.MoveFirst
    StoreMemvars
End Sub

Private Sub cmdLast_Click()
    RSORD_HD.MoveLast
    StoreMemvars
End Sub

Private Sub cmdNext_Click()
    RSORD_HD.MoveNext
    If RSORD_HD.EOF Then
        RSORD_HD.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPISNum_Click()
    Set FRM_SERIES = New frmPMISTrans_ADB_Issuances_PISFormation

    With FRM_SERIES
        .SetStockType ("A")
        If ADDOREDIT = "EDIT" Then
            .EditSeries (txtReferencePIS)
            .txtedit = "EDIT"
        Else
            .txtedit = ""
        End If
        .lbl2 = Mid(txtReferencePIS, 3, 1)
        .lbl3 = Mid(txtReferencePIS, 4, 1)
        .lbl4 = Mid(txtReferencePIS, 5, 1)
        .lbl9.Text = Mid(txtReferencePIS, 9, 3)
        .lbl11 = Mid(txtReferencePIS, 12, 1)
        If .lbl2.Caption = "S" Then
            .optS.Value = True
        ElseIf .lbl2.Caption = "W" Then
            .optW.Value = True
        ElseIf .lbl2.Caption = "M" Then
            .optM.Value = True
        ElseIf .lbl2.Caption = "J" Then
            .optJ.Value = True
        ElseIf .lbl2.Caption = "O" Then
            .optO.Value = True
        End If
        If .lbl3.Caption = "G" Then
            .optG.Value = True
        ElseIf .lbl3.Caption = "B" Then
            .optB.Value = True
        End If
        If .lbl4.Caption = "C" Then
            .optC.Value = True
        ElseIf .lbl4.Caption = "I" Then
            .optI.Value = True
        ElseIf .lbl4.Caption = "W" Then
            .optW2.Value = True
        End If
        If .lbl11.Caption = "1" Then
            .opt1.Value = True
        ElseIf .lbl11.Caption = "2" Then
            .opt2.Value = True
        ElseIf .lbl11.Caption = "0" Then
            .opt0.Value = True
        End If
    End With
    FRM_SERIES.Show 1
    On Error Resume Next
    txtCustName.SetFocus
End Sub

Private Sub Command1_Click()
    otp_ADB_Click
    pic_viewStockADB.ZOrder 0
    pic_viewStockADB.Visible = True
    TXT_ROVIEW = txtRONO
    ShowStock_Issused txtRONO
End Sub

Private Sub frm_series_SETSERIES(XDATE As String, XSERIES As String)
    txtReferencePIS = XSERIES
    txtTranDate = XDATE
End Sub

Private Sub otp_RIVADB_Click()
    AddColumnHeader "Stock#, Onhand,ADB Qty,Isssuances To ADB,Balance", ListView2
    ResizeColumnHeader ListView2, "15,20,20,20,20"
    ListView2.ListItems.Clear
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
End Sub

Private Sub otp_ADB_Click()
    AddColumnHeader "Tranno,Date,Stock#, Description,Onhand,ADB Qty", ListView2
    ResizeColumnHeader ListView2, "8,12,16,35,10,10"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ListItems.Clear
End Sub

Sub ShowStock_Issused(VTXTRONO As String)
    Dim RSADB                                          As ADODB.Recordset
    Dim LST                                            As ListItem
    Dim issused_qty                                    As Long
    Dim balnces                                        As Long
    ListView2.ListItems.Clear
    If otp_RIVADB.Value = True Then
        Set RSADB = gconDMIS.Execute("SELECT STOCK_ORD ,AVG(PMIS_STOCKMAS.ONHAND) ONHAND , sum(TRANQTY) as TRANQTY FROM PMIS_ALLDAYTRAN INNER JOIN PMIS_STOCKMAS " & _
                                   " ON PMIS_STOCKMAS.TYPE=PMIS_ALLDAYTRAN.TYPE AND PMIS_ALLDAYTRAN.STOCK_ORD=PMIS_STOCKMAS.STOCKNO " & _
                                   " WHERE  PMIS_STOCKMAS.TYPE='M' AND " & _
                                   " (TRANNO  IN (SELECT TRANNO FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='M' AND RONO='" & VTXTRONO & "'   AND (STATUS='P' OR STATUS='B')) " & _
                                   " OR TRANNO IN (SELECT TRANNO FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='M' AND RONO='" & VTXTRONO & "'   AND (STATUS='P' OR STATUS='B'))) " & _
                                   " AND TRANTYPE='ADB' GROUP BY STOCK_ORD  ORDER BY STOCK_ORD")

        While Not RSADB.EOF
            Set LST = ListView2.ListItems.Add(, , Null2String(RSADB!STOCK_ORD))
            Call LST.ListSubItems.Add(, , Null2String(RSADB!ONHAND))
            Call LST.ListSubItems.Add(, , Null2String(RSADB!TRANQTY))
            issused_qty = GetTotal_ADB_Filled(VTXTRONO, Null2String(RSADB!STOCK_ORD))
            Call LST.ListSubItems.Add(, , issused_qty)
            balnces = N2Str2Zero(RSADB!TRANQTY) - issused_qty
            Call LST.ListSubItems.Add(, , balnces)

            RSADB.MoveNext
        Wend


    Else
        Set RSADB = gconDMIS.Execute("SELECT PMIS_ALLDAYTRAN.TRANNO, PMIS_ALLDAYTRAN.TRANDATE, STOCK_ORD ,PMIS_STOCKMAS.STOCKDESC ,PMIS_STOCKMAS.ONHAND , TRANQTY FROM PMIS_ALLDAYTRAN INNER JOIN PMIS_STOCKMAS " & _
                                   " ON PMIS_STOCKMAS.TYPE=PMIS_ALLDAYTRAN.TYPE AND PMIS_ALLDAYTRAN.STOCK_ORD=PMIS_STOCKMAS.STOCKNO WHERE PMIS_STOCKMAS.TYPE='M' AND " & _
                                   " (TRANNO  IN (SELECT TRANNO FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='M' AND RONO='" & VTXTRONO & "'   AND (STATUS='P' OR STATUS='B')) " & _
                                   " OR TRANNO IN (SELECT TRANNO FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='M' AND RONO='" & VTXTRONO & "'   AND (STATUS='P' OR STATUS='B'))) " & _
                                   " AND TRANTYPE='ADB' ORDER BY STOCK_ORD")

        While Not RSADB.EOF
            Set LST = ListView2.ListItems.Add(, , Null2String(RSADB!TRANNO))
            Call LST.ListSubItems.Add(, , Null2String(RSADB!trandate))
            Call LST.ListSubItems.Add(, , Null2String(RSADB!STOCK_ORD))
            Call LST.ListSubItems.Add(, , Null2String(RSADB!STOCKDESC))
            Call LST.ListSubItems.Add(, , Null2String(RSADB!ONHAND))
            Call LST.ListSubItems.Add(, , Null2String(RSADB!TRANQTY))
            'Call lst.ListSubItems.Add(, , GetTotal_ADB_Filled(vtxtRONO, Null2String(RSADB!STOCK_ORD)))
            RSADB.MoveNext
        Wend
    End If





End Sub
Private Sub Command4_Click()
    pic_viewStockADB.Visible = False
End Sub

Private Sub Command5_Click()
    ShowStock_Issused Replace(TXT_ROVIEW, "'", "")
End Sub

Private Sub Command6_Click()
    pic_viewStockADB.ZOrder 0
    pic_viewStockADB.Visible = True
End Sub

Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub

    txtRONO = ListView1.SelectedItem.ListSubItems(2).Text
    Dim RONOStr                                        As String
    RONOStr = txtRONO.Text
    If Left(RONOStr, 2) = "R-" Then
        RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
    Else
        RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
    End If
    txtRONO.Text = RONOStr
    SetCustInfo (RONOStr)
    cmdClose_1_Click
End Sub




Private Sub Text1_Change()
    Dim RSLIST                                         As ADODB.Recordset
    Dim SQLX                                           As String
    Dim str_RONO                                       As String
    Dim searchstring                                   As String

    If LTrim(RTrim(Text1)) <> "" Then
        str_RONO = Text1.Text
        If Left(str_RONO, 2) = "R-" Then
            str_RONO = "R-" & Format(NumericVal(Right(str_RONO, Len(str_RONO) - 2)), "00000000")
        Else
            str_RONO = "R-" & Format(NumericVal(Right(str_RONO, Len(str_RONO))), "00000000")
        End If
        searchstring = " AND  rono like " & N2Str2Null(str_RONO & "%")
    End If

    SQLX = "SELECT TRANDATE,TRANNO ,RONO ,'CURRT' AS DSTATUS FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='M' AND ISNULL(STATUS3,'')  <>'F' AND ISNULL(STATUS2,'')  <>'R' AND (STATUS='P' OR STATUS='B') " & searchstring & vbCrLf
    SQLX = SQLX & " UNION " & vbCrLf
    SQLX = SQLX & "SELECT TRANDATE,TRANNO ,RONO ,'HIST' AS DSTATUS FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='M' AND ISNULL(STATUS3,'')  <>'F' AND  ISNULL(STATUS2,'')  <>'R' AND (STATUS='P' OR STATUS='B') " & searchstring & " ORDER BY RONO"
    Set RSLIST = gconDMIS.Execute(SQLX)
    Listview_Loadval ListView1.ListItems, RSLIST
End Sub

Private Sub TXT_ROVIEW_LostFocus()
    Dim RONOStr                                        As String
    RONOStr = TXT_ROVIEW
    If Left(RONOStr, 2) = "R-" Then
        RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
    Else
        RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
    End If
    TXT_ROVIEW = RONOStr
End Sub

Private Sub txtPRtranno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTranNo.Text = txtPRtranno
        txtPRtranno.Visible = False
    End If
End Sub

Private Sub cmdPost_Click()
    Dim rsPrtMas                                       As New ADODB.Recordset
    Dim rsTdytran                                      As New ADODB.Recordset
    Dim blnStockremove                                 As Boolean
    blnStockremove = False
    If Function_Access(LOGID, "Acess_Post", LOCALACESS) = False Then Exit Sub

    Dim fild                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    fild = grdDetails.Text
    If fild = "" Or fild = "No Entry" Then
        MsgBox "Posting of Transaction without issuance of Line Item(s) is not allowed.", vbCritical, "Pls. Add Part(s)."
        Exit Sub
    End If
    rsTdytran.Open ("SELECT STOCK_ORD,TRANQTY, ID FROM PMIS_TDAYTRAN WHERE TRANNO = '" & txtTranNo & "' AND TYPE='A' AND TRANTYPE IN('RIV') "), gconDMIS
    If Not (rsTdytran.BOF And rsTdytran.EOF) Then
        Do While Not rsTdytran.EOF
            rsPrtMas.Open "SELECT STOCKNO,ONHAND FROM PMIS_STOCKMAS WHERE STOCKNO = '" & rsTdytran!STOCK_ORD & "' AND TYPE='M'", gconDMIS
            '=[ EAP:040209: this will remove the partnumber without stock in the transaction. ]=
            If Not (rsPrtMas.BOF And rsPrtMas.EOF) Then
                If rsPrtMas!ONHAND <= 0 Then
                    MsgBox "Stock# " & rsTdytran!STOCK_ORD & " will be remove from the transaction Out of Stock"
                    SQL_STATEMENT = "delete from PMIS_TdayTran where Id = '" & rsTdytran!ID & "' "
                    gconDMIS.Execute SQL_STATEMENT
                    blnStockremove = True
                ElseIf rsPrtMas!ONHAND < rsTdytran!TRANQTY Then
                    MsgBox "Some Stock# Onhand Is Less that your requested quantity.", vbInformation
                    Exit Sub
                End If
                rsPrtMas.MoveNext
            End If
            rsPrtMas.Close

            rsTdytran.MoveNext
        Loop
    End If

    If blnStockremove Then
        cmdTranCancel.Value = True
        rsRefresh
        Exit Sub
    End If

    If MsgQuestionBox("Are you sure you want to Post this Transaction?", "Post Transaction") = False Then: Exit Sub


    Dim PCURONHAND                                     As Long
    Dim PCurTISSQTY                                    As Long
    Dim PCURISSUANCES                                  As Long
    Dim rsTdaytranDup                                  As ADODB.Recordset
    Dim rsPartmasDup                                   As ADODB.Recordset

    Set rsTdaytranDup = New ADODB.Recordset
    rsTdaytranDup.Open "SELECT ID,TRANTYPE,TRANNO,STOCK_ORD,TRANQTY FROM PMIS_TDAYTRAN WHERE TYPE='A' AND TRANNO = " & N2Str2Null(RSORD_HD!TRANNO) & " AND TRANTYPE = " & N2Str2Null(RSORD_HD!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
        rsTdaytranDup.MoveFirst
        Do While Not rsTdaytranDup.EOF
            Set rsPartmasDup = New ADODB.Recordset
            rsPartmasDup.Open "SELECT STOCKNO,ONHAND,TISSQTY,ISSUANCES,REQSERVED,S_REQSERVED,NON_HARI FROM PMIS_STOCKMAS WHERE STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD) & " AND TYPE='M'", gconDMIS

            If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                PCURONHAND = N2Str2IntZero(rsPartmasDup!ONHAND) - N2Str2Zero(rsTdaytranDup!TRANQTY)
                PCurTISSQTY = N2Str2IntZero(rsPartmasDup!TISSQTY) + N2Str2Zero(rsTdaytranDup!TRANQTY)
                PCURISSUANCES = N2Str2IntZero(rsPartmasDup!ISSUANCES) + N2Str2Zero(rsTdaytranDup!TRANQTY)

                If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                    SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                  " REQSERVED = " & N2Str2IntZero(rsPartmasDup!REQSERVED) + N2Str2Zero(rsTdaytranDup!TRANQTY) & _
                                  " WHERE TYPE='M' AND STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                    gconDMIS.Execute SQL_STATEMENT
                    '===================================================================
                    NEW_LogAudit "E", "ACCESSORIES MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""
                    '===================================================================
                Else
                    SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                  " S_REQSERVED = " & N2Str2IntZero(rsPartmasDup!S_REQSERVED) + N2Str2Zero(rsTdaytranDup!TRANQTY) & _
                                  " WHERE TYPE='M' AND STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                    gconDMIS.Execute SQL_STATEMENT
                    '===================================================================
                    NEW_LogAudit "E", "ACCESSORIES MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""
                    '===================================================================
                End If

                SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                              " ONHAND = " & PCURONHAND & "," & _
                              " TISSQTY = " & PCurTISSQTY & "," & _
                              " ISSUANCES = " & PCURISSUANCES & "," & _
                              " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                              " LASTUPDATE = '" & LOGDATE & "'" & _
                              " WHERE TYPE='M' AND STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                gconDMIS.Execute SQL_STATEMENT
                '===================================================================
                NEW_LogAudit "E", "ACCESSORIES MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""
                '===================================================================
                SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET" & _
                              " STATUS = 'M'," & _
                              " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                              " LASTUPDATE = '" & LOGDATE & "'" & _
                              " WHERE ID = " & rsTdaytranDup!ID
                gconDMIS.Execute SQL_STATEMENT
                '===================================================================
                NEW_LogAudit "PP", LOCALACESS, SQL_STATEMENT, labID, "Parts", txtTranNo, COUNTERTYPE, ""
                '===================================================================

            End If
            rsTdaytranDup.MoveNext
        Loop
    End If
    SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                  " STATUS = 'P'," & _
                  " TOTALQTY = " & ORD_TOTQTY & "," & _
                  " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                  " LASTUPDATE = '" & LOGDATE & "'" & _
                  " WHERE ID = " & labID.Caption

    gconDMIS.Execute SQL_STATEMENT
    '===================================================================
    NEW_LogAudit "P", LOCALACESS, SQL_STATEMENT, labID, "ACCESSORIES", txtTranNo, COUNTERTYPE, ""
    '===================================================================
    rsRefresh
    RSORD_HD.Find "id =" & labID.Caption
    StoreMemvars

    Set rsTdaytranDup = Nothing
    Set rsPartmasDup = Nothing

    Dim RSADB                                          As ADODB.Recordset
    Dim STR_SQLX                                       As String
    Dim BLN_ADB_STATUS                                 As Boolean

    BLN_ADB_STATUS = False

    STR_SQLX = " SELECT STOCK_ORD,SUM(TRANQTY) AS  TRANQTY  FROM PMIS_TDAYTRAN "
    STR_SQLX = STR_SQLX & " INNER JOIN  PMIS_ORD_HD ON "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HD.TYPE=PMIS_TDAYTRAN.TYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HD.TRANTYPE=PMIS_TDAYTRAN.TRANTYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HD.TRANNO = PMIS_TDAYTRAN.TRANNO "
    STR_SQLX = STR_SQLX & " WHERE PMIS_ORD_HD.TRANTYPE='ADB' AND  PMIS_ORD_HD.RONO='" & txtRONO & "' AND "
    STR_SQLX = STR_SQLX & " (PMIS_ORD_HD.STATUS='P' OR PMIS_ORD_HD.STATUS='B') AND PMIS_ORD_HD.TYPE='M' "
    STR_SQLX = STR_SQLX & " GROUP BY STOCK_ORD"
    STR_SQLX = STR_SQLX & " Union "

    STR_SQLX = STR_SQLX & " SELECT STOCK_ORD,SUM(TRANQTY) AS  TRANQTY  FROM PMIS_DAYTRAN "
    STR_SQLX = STR_SQLX & " INNER JOIN  PMIS_ORD_HIST ON "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TYPE=PMIS_DAYTRAN.TYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TRANTYPE=PMIS_DAYTRAN.TRANTYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TRANNO = PMIS_DAYTRAN.TRANNO "
    STR_SQLX = STR_SQLX & " WHERE PMIS_ORD_HIST.TRANTYPE='ADB' AND  PMIS_ORD_HIST.RONO='" & txtRONO & "' AND "
    STR_SQLX = STR_SQLX & " (PMIS_ORD_HIST.STATUS='P' OR PMIS_ORD_HIST.STATUS='B')  AND PMIS_ORD_HIST.TYPE='M' "
    STR_SQLX = STR_SQLX & " GROUP BY STOCK_ORD"

    Set RSADB = gconDMIS.Execute(STR_SQLX)
    Dim FILLED_QTY                                     As Long
    Dim ADB_QTY                                        As Long
    While Not RSADB.EOF
        ADB_QTY = RSADB!TRANQTY
        FILLED_QTY = GetTotal_ADB_Filled(txtRONO, Null2String(RSADB!STOCK_ORD))

        If ADB_QTY > FILLED_QTY Then
            BLN_ADB_STATUS = True
        End If
        RSADB.MoveNext
    Wend


    If BLN_ADB_STATUS = False Then
        gconDMIS.Execute ("UPDATE PMIS_ORD_HD SET STATUS3='F' WHERE RONO=" & N2Str2Null(txtRONO))
        gconDMIS.Execute ("UPDATE PMIS_ORD_HIST SET STATUS3='F' WHERE RONO=" & N2Str2Null(txtRONO))
    
'     Dim LNG As Long
'        Dim RSTDAYTRAN As ADODB.Recordset
'        If CheckIfROBilled(txtRONO) = "" Then
'        Set RSTDAYTRAN = gconDMIS.Execute("SELECT STOCK_ORD, MAC FROM PMIS_DAYTRAN WHERE TRANNO=" & N2Str2Null(txtTranNo) & " AND TYPE='M' AND TRANTYPE='RIV'")
'            While Not RSTDAYTRAN.EOF
'                Call gconDMIS.Execute("UPDATE CSMS_RO_DET SET DETCOST =" & N2Str2Zero(RSTDAYTRAN!Mac) & " WHERE LIVIL='3' AND DETCDE=" & N2Str2Null(RSTDAYTRAN!STOCK_ORD), LNG)
'                RSTDAYTRAN.MoveNext
'            Wend
'       End If
        
    End If
    

    Exit Sub
Errorcode:
    ShowVBError
End Sub
Function GetTotal_ADB_Filled(xro_no As String, x_stockno As String) As Long
    Dim STR_SQLX                                       As String

    STR_SQLX = " SELECT STOCK_ORD,SUM(TRANQTY) AS  TRANQTY  FROM PMIS_TDAYTRAN "
    STR_SQLX = STR_SQLX & " INNER JOIN  PMIS_ORD_HD ON "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HD.TYPE=PMIS_TDAYTRAN.TYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HD.TRANTYPE=PMIS_TDAYTRAN.TRANTYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HD.TRANNO = PMIS_TDAYTRAN.TRANNO "
    STR_SQLX = STR_SQLX & " WHERE PMIS_ORD_HD.TRANTYPE='RIV' AND  PMIS_ORD_HD.RONO='" & xro_no & "' AND PMIS_TDAYTRAN.STOCK_ORD='" & x_stockno & "' AND "
    STR_SQLX = STR_SQLX & " (PMIS_ORD_HD.STATUS='P' OR PMIS_ORD_HD.STATUS='B')  AND PMIS_ORD_HD.TYPE='M' AND PMIS_ORD_HD.STATUS2='R' "
    STR_SQLX = STR_SQLX & " GROUP BY STOCK_ORD"
    STR_SQLX = STR_SQLX & " Union "

    STR_SQLX = STR_SQLX & " SELECT STOCK_ORD,SUM(TRANQTY) AS  TRANQTY  FROM PMIS_DAYTRAN "
    STR_SQLX = STR_SQLX & " INNER JOIN  PMIS_ORD_HIST ON "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TYPE=PMIS_DAYTRAN.TYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TRANTYPE=PMIS_DAYTRAN.TRANTYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TRANNO = PMIS_DAYTRAN.TRANNO "
    STR_SQLX = STR_SQLX & " WHERE PMIS_ORD_HIST.TRANTYPE='RIV' AND  PMIS_ORD_HIST.RONO='" & xro_no & "' AND PMIS_DAYTRAN.STOCK_ORD='" & x_stockno & "' AND "
    STR_SQLX = STR_SQLX & " (PMIS_ORD_HIST.STATUS='P' OR PMIS_ORD_HIST.STATUS='B')  AND PMIS_ORD_HIST.TYPE='M' AND PMIS_ORD_HIST.STATUS2='R'"
    STR_SQLX = STR_SQLX & " GROUP BY STOCK_ORD"

    Dim RSTOTAL_FILLED                                 As ADODB.Recordset
    Set RSTOTAL_FILLED = gconDMIS.Execute(STR_SQLX)
    If Not RSTOTAL_FILLED.EOF Or Not RSTOTAL_FILLED.BOF Then
        GetTotal_ADB_Filled = N2Str2Zero(RSTOTAL_FILLED!TRANQTY)
    End If




End Function
Private Sub cmdPrevious_Click()
    RSORD_HD.MovePrevious
    If RSORD_HD.BOF Then
        RSORD_HD.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", LOCALACESS) = False Then Exit Sub


    If MsgQuestionBox("Parts Issuance Slip will be printed. You want to print it in a Blank form?", "Confirm Printing...") = True Then
        If COMPANY_CODE = "HCI" Then
            cmdPrintRIV_Click
        Else
            fraSignatories.Visible = True
            fraSignatories.ZOrder 0
            txtPreparedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", "")
            txtIssuedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", "")
            txtApprovedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", "")
            On Error Resume Next
            txtRequestedBy.SetFocus
        End If
    Else
        SERVICEPISPRINTING
    End If


    NEW_LogAudit "V", LOCALACESS, "", labID, "Parts", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdPrintRIV_Click()
    SERVICEPISPRINTING_BLANKFORM
    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", txtPreparedBy)
    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", txtIssuedBy)
    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", txtApprovedBy)

    SendToBack
End Sub

Private Sub cmdSave_Click()
    'JAA
    If IsDate(txtTranDate) = True Then
        If DateDiff("m", txtTranDate, LOGDATE) <> 0 Then
            MsgBox "Warning: Transaction Month cannot be greater or less than the current month.", vbCritical
            On Error Resume Next
            txtTranDate.SetFocus
        End If
    Else
        MsgBox "Invalid Date", vbInformation
        Exit Sub
    End If

    If gconDMIS.Execute("SELECT COUNT(*) FROM CSMS_REPOR WHERE REP_OR=" & N2Str2Null(LTrim(RTrim(Replace(txtRONO, "'", ""))))).Fields(0).Value = 0 Then
        MsgBox "RO Number Doesn't Exists. Please Correct Repair Order Number", vbInformation
        On Error Resume Next
        txtRONO.SetFocus
        Exit Sub
    End If




    On Error GoTo Errorcode
    Dim NEXTCUNTER                                     As String
    Dim RSFINDDUP                                      As ADODB.Recordset
    Dim XSALES_ORIGIN                                  As String
    Dim XSI_TYPE                                       As String
    Dim XPAY_CLASS                                     As String
    Dim XCHAR_YEAR                                     As String
    Dim XCHAR_MONTH                                    As String
    Dim XIS_SERIES                                     As String
    Dim XTRACK_CODE                                    As String
    Dim VCBOSALESMAN                                   As String
    Dim VCBOSMNAME                                     As String
    Dim VTXTTRANTYPE                                   As String
    Dim VTXTTRANNO                                     As String
    Dim VTXTTRANDATE                                   As String
    Dim VTXTCUSTCODE                                   As String
    Dim VTXTCUSTNAME                                   As String
    Dim VTXTCHARGETO                                   As String
    Dim VTXTREP_OR                                     As String
    Dim VTXTREFPRSNO                                   As String
    Dim VTXTRONO                                       As String
    Dim VtxtTerms                                      As String
    Dim VStatus                                        As String
    Dim VTXTTTLINVAMT                                  As Double
    Dim VTXTDS1                                        As Double
    Dim VTXTDS_Desc1                                   As String
    Dim VTXTDS_Amt1                                    As Double
    Dim VTXTNETINVAMT                                  As Double
    Dim VTXTRemarks                                    As String
    Dim Vusercode                                      As String
    Dim VLastUpdate                                    As String
    Dim VIN_PROCESS                                    As String
    Dim VTXTREFERENCEPIS                               As String

    If Len(Trim(RTrim(txtTranNo))) <> 6 Then
        MsgBox "Invalid Transaction Number. Should Be Six Digit In Length!", vbCritical, "Transaction Number!"
        On Error Resume Next
        txtTranNo.SetFocus
        Exit Sub
    End If

    If Trim(txtReferencePIS.Text) = "" Or Len(txtReferencePIS.Text) < 10 Then
        MsgBox "Invalid Reference PIS Number!", vbCritical, "PIS Required!"
        Exit Sub
    End If

    If Trim(txtRONO.Text) = "" Then
        MsgBox "RO Number is Required...", vbInformation, "Pls Input RO Number..."
        Exit Sub
    End If

    If RTrim(LTrim(cboRefPRSNo.Text)) = "" Then
        MsgBox "Reference PRS Number is Required...", vbInformation, "Pls. select PRS No."
        'Exit Sub
    End If

    If IsNull(txtTranNo.Text) = True Then
        MsgSpeechBox "Transaction No. must not be empty"
        On Error Resume Next
        txtTranNo.SetFocus
        Exit Sub
    Else
        If ADDOREDIT = "ADD" Then
            Set RSFINDDUP = New ADODB.Recordset

            RSFINDDUP.Open "SELECT TRANTYPE,TRANNO FROM PMIS_VW_ISS_HISTORY WHERE TYPE='A' AND TRANTYPE = '" & txtTranType.Text & "' AND TRANNO = '" & txtTranNo.Text & "' ORDER BY TRANTYPE,TRANNO", gconDMIS, adOpenForwardOnly, adLockReadOnly

            If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                MsgSpeechBox "Transaction No. already exist!"
                txtTranNo.SetFocus
                On Error Resume Next
                Exit Sub
            End If
        Else
            If LTrim(RTrim(txtTranNo)) <> LTrim(RTrim(Null2String(RSORD_HD!TRANNO))) Then
                Set RSFINDDUP = New ADODB.Recordset
                RSFINDDUP.Open "SELECT TRANTYPE,TRANNO FROM PMIS_VW_ISS_HISTORY WHERE TYPE='A' AND TRANTYPE = '" & txtTranType.Text & "' AND TRANNO = '" & txtTranNo.Text & "' ORDER BY TRANTYPE,TRANNO", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                    MsgSpeechBox "Transaction No. already exist!"
                    On Error Resume Next
                    txtTranNo.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

    If txtTranDate.Text = "" Or IsDate(txtTranDate.Text) = False Then
        MsgSpeechBox "Invalid Transaction Date!"
        On Error Resume Next
        txtTranDate.SetFocus
        Exit Sub
    End If



    VCBOSALESMAN = N2Str2Null(cboSalesMan.Text)
    VCBOSMNAME = N2Str2Null(cboSMName.Text)

    If Left(txtTranNo.Text, 1) = "M" Then
        'do notning for alpha numeric series
    Else
        NEXTCUNTER = NumericVal(txtTranNo.Text) + 1
    End If

    VTXTTRANTYPE = N2Str2Null(txtTranType.Text)
    VTXTTRANNO = N2Str2Null(txtTranNo.Text)
    VTXTTRANDATE = N2Date2Null(txtTranDate.Text)
    VTXTCUSTCODE = N2Str2Null(txtCustCode.Text)
    VTXTCUSTNAME = N2Str2Null(txtCustName.Text)
    VTXTREFERENCEPIS = N2Str2Null(txtReferencePIS.Text)
    VTXTREFPRSNO = N2Str2Null(cboRefPRSNo.Text)
    VIN_PROCESS = "'Y'"
    VTXTCHARGETO = "'VAR'"

    VTXTRONO = N2Str2Null(txtRONO.Text)

    If Len(txtRONO.Text) = 7 Then
        VTXTREP_OR = "'" & Left(txtRONO.Text, 1) & "-" & Right(txtRONO.Text, 6) & "'"
    Else
        VTXTREP_OR = "NULL"
    End If

    VtxtTerms = N2Str2Null(txtTerms.Text)
    VTXTTTLINVAMT = NumericVal(txtTTLInvAmt.Text)
    VTXTDS1 = NumericVal(txtDS1.Text)
    VTXTDS_Desc1 = N2Str2Null(txtDS_Desc1.Text)
    VTXTDS_Amt1 = NumericVal(txtDS_Amt1.Text)
    VTXTNETINVAMT = NumericVal(txtNetInvAmt.Text)
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    XSALES_ORIGIN = N2Str2Null(Mid(txtReferencePIS, 3, 1))
    XSI_TYPE = N2Str2Null(Mid(txtReferencePIS, 4, 1))
    XPAY_CLASS = N2Str2Null(Mid(txtReferencePIS, 5, 1))
    XCHAR_YEAR = N2Str2Null(Mid(txtReferencePIS, 6, 2))
    XCHAR_MONTH = N2Str2Null(Mid(txtReferencePIS, 8, 1))
    XIS_SERIES = N2Str2Null(Mid(txtReferencePIS, 9, 3))
    XTRACK_CODE = N2Str2Null(Mid(txtReferencePIS, 12, 1))
    VStatus = "'N'"

    If txtRemarks.Text = "Pls Type Your Message Here!" Then
        VTXTRemarks = "NULL"
    Else
        VTXTRemarks = Replace(txtRemarks.Text, Chr(13), "")
        VTXTRemarks = Replace(txtRemarks.Text, Chr(9), "")
        VTXTRemarks = Replace(Trim(txtRemarks.Text), Chr(27), "")
        VTXTRemarks = N2Str2Null(VTXTRemarks)
    End If

    If ADDOREDIT = "ADD" Then
        SQL_STATEMENT = "INSERT INTO PMIS_ORD_HD" & _
                      " (STATUS2,TYPE,TRANTYPE,TRANNO,TRANDATE,CUSTCODE,CUSTNAME,CHARGETO,REFPRSNO,RONO,REP_OR,SALESMAN,SMNAME,TERMS,TTLINVAMT,DS1,DS_DESC1,DS_AMT1,NETINVAMT,REMARKS,STATUS,USERCODE,LASTUPDATE,IN_PROCESS,REFPISNO,SALES_ORIGIN,SI_TYPE,PAY_CLASS,CHAR_YEAR,CHAR_MONTH,IS_SERIES,TRACK_CODE)" & _
                      " VALUES ('R','M'," & VTXTTRANTYPE & ", " & VTXTTRANNO & ", " & VTXTTRANDATE & ", " & _
                      " " & VTXTCUSTCODE & ", " & VTXTCUSTNAME & ", " & VTXTCHARGETO & "," & VTXTREFPRSNO & _
                        ", " & VTXTRONO & "," & VTXTREP_OR & ", " & VCBOSALESMAN & ", " & VCBOSMNAME & _
                        ", " & VtxtTerms & ", " & VTXTTTLINVAMT & _
                        ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                        ", " & VTXTNETINVAMT & ", " & VTXTRemarks & _
                        ", " & VStatus & ", " & Vusercode & ", " & VLastUpdate & "," & VIN_PROCESS & "," & VTXTREFERENCEPIS & ", " & XSALES_ORIGIN & ", " & XSI_TYPE & ", " & XPAY_CLASS & ", " & XCHAR_YEAR & ", " & XCHAR_MONTH & ", " & XIS_SERIES & ", " & XTRACK_CODE & ")"

        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("A", LOCALACESS, SQL_STATEMENT, FindTransactionID(txtTranNo, "tranno", "PMIS_Ord_Hd", "DETAILS", N2Str2Null("P"), "TYPE"), "Parts", txtTranNo & " - " & VTXTREFPRSNO, COUNTERTYPE, "")
        ShowSuccessFullyAdded
    Else
        Dim LNG                                        As Long
        If Null2String(RSORD_HD!RoNo) <> LTrim(RTrim(txtRONO)) Then
            LNG = gconDMIS.Execute("SELECT COUNT(*) FROM PMIS_TDAYTRAN WHERE TRANTYPE='RIV' AND TYPE='M' AND TRANNO=" & N2Str2Null(RSORD_HD!TRANNO)).Fields(0).Value
            If LNG > 0 Then
                MsgBox "Editing of RO Not Allowed for Item with Line Item(s)" & vbCrLf & "Please Remove all the list item before Editing RO Number", vbInformation
                cmdCancel_Click
                Exit Sub
            End If
        End If

        SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                      " TRANTYPE = " & VTXTTRANTYPE & "," & _
                      " TRANNO = " & VTXTTRANNO & "," & _
                      " TRANDATE = " & VTXTTRANDATE & "," & _
                      " CUSTCODE = " & VTXTCUSTCODE & "," & _
                      " CUSTNAME = " & VTXTCUSTNAME & "," & _
                      " CHARGETO = " & VTXTCHARGETO & "," & _
                      " REFPRSNO = " & VTXTREFPRSNO & "," & _
                      " RONO = " & VTXTRONO & "," & _
                      " REP_OR = " & VTXTREP_OR & "," & _
                      " SALESMAN = " & VCBOSALESMAN & "," & _
                      " SMNAME = " & VCBOSMNAME & "," & _
                      " TERMS = " & VtxtTerms & "," & _
                      " TTLINVAMT = " & VTXTTTLINVAMT & "," & _
                      " DS1 = " & VTXTDS1 & "," & _
                      " DS_DESC1 = " & VTXTDS_Desc1 & "," & _
                      " DS_AMT1 = " & VTXTDS_Amt1 & "," & _
                      " NETINVAMT = " & VTXTNETINVAMT & "," & _
                      " REMARKS = " & VTXTRemarks & ", " & _
                      " STATUS = " & VStatus & ", " & _
                      " USERCODE = " & Vusercode & ", " & _
                      " IN_PROCESS = " & VIN_PROCESS & ", " & _
                      " REFPISNO = " & VTXTREFERENCEPIS & ", " & _
                      " LASTUPDATE = " & VLastUpdate & _
                      " WHERE ID = " & labID.Caption

        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, labID, "ACCESSORIES", "TRAN NO: " & txtTranNo, COUNTERTYPE, "")

        SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                      " SALES_ORIGIN = " & XSALES_ORIGIN & "," & _
                      " SI_TYPE = " & XSI_TYPE & "," & _
                      " PAY_CLASS = " & XPAY_CLASS & "," & _
                      " CHAR_YEAR = " & XCHAR_YEAR & "," & _
                      " CHAR_MONTH = " & XCHAR_MONTH & "," & _
                      " IS_SERIES = " & XIS_SERIES & "," & _
                      " TRACK_CODE = " & XTRACK_CODE & "" & _
                      " WHERE ID = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, labID, "ACCESSORIES", "TRAN NO: " & txtTranNo & " - " & VTXTREFPRSNO, COUNTERTYPE, "")


        SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET" & _
                      " TRANTYPE = " & VTXTTRANTYPE & "," & _
                      " TRANDATE = " & VTXTTRANDATE & "," & _
                      " TRANNO = " & VTXTTRANNO & _
                      " WHERE TYPE='A' AND TRANTYPE = '" & PREVORDTYPE & "' AND TRANNO = '" & Null2String(RSORD_HD!TRANNO) & "'"
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("EE", LOCALACESS, SQL_STATEMENT, labID, "ACCESSORIES", "TRAN NO: " & txtTranNo & " - " & VTXTREFPRSNO, COUNTERTYPE, "")
        ShowSuccessFullyUpdated
    End If

    If ADDOREDIT = "ADD" Then
        If Left(txtTranNo.Text, 1) = "M" Then
            'DO NOTHING FOR ALPHA NUMERIC SERIES
        Else
            SQL_STATEMENT = "UPDATE PMIS_COUNTER SET NEXTNUMBER = '" & NEXTCUNTER & "', LASTUPDATE = '" & LOGDATE & "', USERCODE = '" & "USER" & "' WHERE TYPE='A' AND MODUL = " & VTXTTRANTYPE
            gconDMIS.Execute SQL_STATEMENT
        End If
        Call NEW_LogAudit("E", "ACCESSORIES COUNTER", SQL_STATEMENT, FindTransactionID(VTXTTRANTYPE, "MODUL", "PMIS_Counter", "DETAILS", N2Str2Null("P"), "TYPE"), "", "MODUL: " & Null2String(VTXTTRANTYPE), "", "")
        X_FillSearchGrid ""
    Else
        SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                      " TTLINVAMT = " & ORD_TOTUPRICE & "," & _
                      " NETINVAMT = " & ORD_TOTINVAMT & _
                      " WHERE TYPE='A' AND TRANNO = " & VTXTTRANNO & " AND TRANTYPE = " & VTXTTRANTYPE
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, labID, "", "TRAN NO: " & txtTranNo, "", "")
    End If

    fraDetails.Enabled = True
    rsRefresh
    RSORD_HD.Find "TRANNO = " & VTXTTRANNO
    cmdCancel.Value = True
    cleargrid grdDetails
    FillDetails
    If ADDOREDIT = "ADD" Then
    cmdAddTran_Click
    End If
    X_FillSearchGrid ""
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdTranCancel_Click()
    SendToBack
    StoreMemvars
End Sub

Private Sub cmdTranDelete_Click()
    On Error GoTo Errorcode:

    If labDetID.Caption = "" Then
        ShowNothingToDeleteMsg
        Exit Sub
    End If

    If MsgQuestionBox("Delete This Line Item, Are you Sure?", "Delete Parts Entry") = True Then
        SQL_STATEMENT = "DELETE FROM PMIS_TDAYTRAN WHERE ID = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", LOCALACESS, SQL_STATEMENT, labID, "ACCESSORIES", "ACCESSORIES NO: " & cboTranPartNo, COUNTERTYPE, labDetID
        ShowDeletedMsg
    End If

    Dim cnt                                            As Integer
    Dim rsTdaytranDup                                  As ADODB.Recordset
    Set rsTdaytranDup = New ADODB.Recordset
    rsTdaytranDup.Open "SELECT ID,ITEMNO FROM PMIS_TDAYTRAN WHERE TYPE='A' AND TRANTYPE = " & N2Str2Null(COUNTERTYPE) & " AND TRANNO = " & N2Str2Null(RSORD_HD!TRANNO) & " ORDER BY ITEMNO ASC", gconDMIS
    If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
        rsTdaytranDup.MoveFirst
        cnt = 0
        Do While Not rsTdaytranDup.EOF
            cnt = cnt + 1
            SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET ITEMNO = " & Format(cnt, "0000") & " WHERE ID = " & rsTdaytranDup!ID
            gconDMIS.Execute SQL_STATEMENT
            rsTdaytranDup.MoveNext
        Loop
    End If
    FillDetails

    SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                  " TTLINVAMT = " & ORD_TOTUPRICE & "," & _
                  " NETINVAMT = " & ORD_TOTINVAMT & _
                  " WHERE ID = " & labID.Caption
    gconDMIS.Execute SQL_STATEMENT
    Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, labID, "ACCESSORIES", "TRAN NO: " & txtTranNo, COUNTERTYPE, "")

    rsRefresh
    On Error Resume Next
    RSORD_HD.Find "id = " & labID.Caption
    cmdTranCancel.Value = True

    Exit Sub
Errorcode:
    ShowVBError
End Sub


Private Sub cmdTranSave_Click()
    Screen.MousePointer = 11
    On Error GoTo Errorcode
    If NumericVal(txtTranQty) <= 0 Then
        MessagePop InfoVoid, "Invalid Input", "Please Input Valid Quantity"
        Exit Sub
    End If

    If NumericVal(txtTranQty) > NumericVal(TXT_AVL_ISSUE) Then
        MessagePop InfoVoid, "Invalid Input", "Over Issuances to Advance Bill"
        Exit Sub
    End If

    If cboTranPartNo.Text = "" Then
        MsgSpeechBox "Warning: Stock Number must have a value"
        On Error Resume Next
        cboTranPartNo.SetFocus
        Exit Sub
    End If

    If ADDOREDIT = "ADD" Then
        Dim rsTDaytranClone                            As ADODB.Recordset
        Set rsTDaytranClone = New ADODB.Recordset
        rsTDaytranClone.Open "SELECT TRANTYPE,TRANNO,ITEMNO,STOCK_ORD FROM PMIS_TDAYTRAN WHERE TYPE='A' AND STOCK_ORD = '" & cboTranPartNo.Text & "' AND TRANTYPE = '" & txtTranType.Text & "' AND TRANNO =" & N2Str2Null(RSORD_HD!TRANNO) & " ORDER BY ITEMNO ASC", gconDMIS
        If Not rsTDaytranClone.EOF And Not rsTDaytranClone.BOF Then
            MsgSpeechBox "Part Number already used in this transaction"
            On Error Resume Next
            cboTranPartNo.SetFocus
            Exit Sub
        End If
        Set rsTDaytranClone = Nothing
    End If

    Dim ORDTRANDATE                                    As String
    Dim ORDTRANNO                                      As String
    Dim ORDTRANTYPE                                    As String
    Dim ORDITEMNO                                      As String
    Dim ORDSTOCK_ORD                                   As String
    Dim ORDSTOCK_SUP                                   As String
    Dim ORDTRANQTY                                     As Long
    Dim ORDTRANUCOST                                   As Double
    Dim ORDSTATUS, ORDIN_OUT                           As String
    Dim ORDTRANINVAMT                                  As Double
    Dim ORDMAC                                         As Double
    Dim CRITICAL_QUESTION                              As String

    If txtTranType.Text <> "ADB" Then
        Dim CURONHAND                                  As Long
        Dim CURSAFESTOCK
        Dim CURTISSQTY                                 As Long
        Dim CURRESSERVICE                              As Long
        Dim CURISSUANCES                               As Long
        Dim PREVCURORDQTY                              As Long
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "SELECT STOCKNO,ONHAND,SSTOCK,RESSERVICE,TISSQTY,ISSUANCES,MAC,NON_HARI FROM PMIS_STOCKMAS WHERE STOCKNO = '" & cboTranPartNo.Text & "' AND ACTIVE = 'Y' AND TYPE='M'", gconDMIS
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            CURONHAND = N2Str2IntZero(RSPARTMAS!ONHAND)
            CURSAFESTOCK = N2Str2IntZero(RSPARTMAS!SSTOCK)
            CURTISSQTY = N2Str2IntZero(RSPARTMAS!TISSQTY)
            CURRESSERVICE = N2Str2IntZero(RSPARTMAS!RESSERVICE)
            CURISSUANCES = N2Str2IntZero(RSPARTMAS!ISSUANCES)
            ORDMAC = NumericVal(RSPARTMAS!Mac)
            If ORDMAC <= 0 Then
                Screen.MousePointer = 0
                MsgBox "Warning: This Stock Number has Zero Cost! Pls Check in Stock Master File or Process Update Master File to Proceed.", vbCritical, "Stock Has Zero Cost"
                Screen.MousePointer = 0
                Exit Sub
            Else
                txtTranUCost.Text = ORDMAC
            End If
            If ADDOREDIT <> "ADD" Then
                PREVCURORDQTY = NumericVal(labPrevOrdQty.Caption)
                CURTISSQTY = CURTISSQTY - PREVCURORDQTY
                CURISSUANCES = CURISSUANCES - PREVCURORDQTY
            End If
            If CURONHAND <= 0 Then
                Screen.MousePointer = 0
                MsgSpeechBox "Out of Stock!"
                Exit Sub
            End If


            If NumericVal(txtTranQty.Text) > CURONHAND Then
                Screen.MousePointer = 0
                MsgSpeechBox "Qty Ordered Exceeds Current Stock!"
                On Error Resume Next
                txtTranQty.SetFocus
                Exit Sub
            Else
                CURONHAND = CURONHAND - NumericVal(txtTranQty.Text)
            End If

            If CURONHAND < CURSAFESTOCK Then
                Screen.MousePointer = 0
                If MsgQuestionBox("Current On-hand is now below the Safety Stock Level... Proceed Anyway?", "Safety Stock Alert!") = False Then
                    Exit Sub
                End If
                CRITICAL_QUESTION = "Current On-hand is now below the Safety Stock Level... Proceed Anyway?"
                Call NEW_LogAudit("MP", LOCALACESS, CRITICAL_QUESTION, labID, "", "TRAN NO: " & txtTranNo & " " & " ACCESSORIES NO: " & cboTranPartNo & " " & CRITICAL_QUESTION, COUNTERTYPE, "")
                MsgBox "User Action has been Log to Audit Trail", vbInformation, "Audit Trail Information"
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
    ORDSTOCK_SUP = N2Str2Null(cboTranPartNo.Text)
    ORDTRANQTY = NumericVal(txtTranQty.Text)
    ORDTRANUCOST = ORDMAC
    ORDTRANINVAMT = NumericVal(txtTranUPrice.Text)

    If Round(ORDTRANINVAMT, 2) < Round(ORDTRANUCOST, 2) Then
        If COMPANY_CODE = "HAS" Then
            If Mid(txtReferencePIS.Text, 3, 1) = "W" Then
                If MsgBox("Issuance Unit Price for this Part Number is less than its Cost!" & vbCrLf & " Do you want to Proceed", vbQuestion + vbYesNo, "PMIS") = vbNo Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                CRITICAL_QUESTION = "Issuance Unit Price for this Part Number is less than its Cost!" & vbCrLf & " Do you want to Proceed"
                Call NEW_LogAudit("MP", LOCALACESS, CRITICAL_QUESTION, labID, "", "ACCESSORIES NO: " & txtTranNo & " " & " ACCESSORIES NO: " & cboTranPartNo & " " & CRITICAL_QUESTION, COUNTERTYPE, "")
                MsgBox "User Action has been Log to Audit Trail", vbInformation, "Audit Trail Information"
            End If
        Else
            Screen.MousePointer = 0
            MsgBox "Warning: Issuance Unit Price for this Part Number is less than its Cost!" & vbCrLf & "                System will not allow this transaction to Proceed.", vbCritical, "Unit Price is Below Cost"
            Exit Sub
        End If
    End If

    ORDIN_OUT = "'O'"
    ORDSTATUS = "'N'"

    Dim HARI_NONHARI                                   As String
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT NON_HARI FROM PMIS_STOCKMAS WHERE TYPE='M' AND STOCKNO = '" & cboTranPartNo.Text & "' AND ACTIVE = 'Y'")
    If Not RSTMP.EOF And Not RSTMP.BOF Then
        HARI_NONHARI = N2Str2Null(RSTMP!NON_HARI)
    Else
        HARI_NONHARI = N2Str2Null("")
    End If
    If ADDOREDIT = "ADD" Then
        SQL_STATEMENT = "INSERT INTO PMIS_TDAYTRAN " & _
                        "(TYPE,TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST,MAC,TRANUPRICE,LASTUPDATE,USERCODE,STATUS,IN_OUT,NON_HARI)" & _
                      " VALUES ('M'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                      " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                      " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                      " " & ORDTRANUCOST & "," & ORDMAC & "," & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & "," & HARI_NONHARI & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "AA", LOCALACESS, SQL_STATEMENT, FindTransactionID(N2Str2Null(txtTranNo), "TRANNO", "PMIS_ORD_HD", "DETAILS", N2Str2Null(Null2String(ORDTRANTYPE)), "TRANTYPE"), "Parts", "ACCESSORIES NO: " & cboTranPartNo, COUNTERTYPE, labDetID
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET" & _
                      " TRANDATE = " & ORDTRANDATE & "," & _
                      " TRANTYPE = " & ORDTRANTYPE & "," & _
                      " TRANNO = " & ORDTRANNO & "," & _
                      " ITEMNO = " & ORDITEMNO & "," & _
                      " STOCK_ORD = " & ORDSTOCK_ORD & "," & _
                      " STOCK_SUP = " & ORDSTOCK_SUP & "," & _
                      " MAC= " & ORDMAC & "," & _
                      " TRANQTY = " & ORDTRANQTY & "," & _
                      " TRANUCOST = " & ORDTRANUCOST & "," & _
                      " TRANUPRICE = " & ORDTRANINVAMT & "," & _
                      " LASTUPDATE = '" & LOGDATE & "'," & _
                      " STATUS = " & ORDSTATUS & "," & _
                      " IN_OUT = " & ORDIN_OUT & "," & _
                      " USERCODE = " & N2Str2Null(LOGCODE) & "" & _
                      " WHERE ID = " & labDetID.Caption

        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", LOCALACESS, SQL_STATEMENT, labID, "ACCESSORIES", "TRAN NO: " & txtTranNo, COUNTERTYPE, labDetID
        ShowSuccessFullyUpdated
    End If
    cleargrid grdDetails
    FillDetails
    SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                  " TOTALQTY = " & ORD_TOTQTY & "," & _
                  " TTLINVAMT = " & ORD_TOTUPRICE & "," & _
                  " NETINVAMT = " & ORD_TOTINVAMT & _
                  " WHERE ID = " & labID.Caption
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT------------------------------------------------------------
    Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, labID, "P", "TRAN NO: " & txtTranNo, "", "")
    'NEW LOG AUDIT------------------------------------------------------------

    Dim rsPRS_Header                                   As ADODB.Recordset
    Dim rsPRS_Details                                  As ADODB.Recordset
    Set rsPRS_Header = gconDMIS.Execute("SELECT * FROM PMIS_VW_PRS WHERE REFPISNO = '" & cboRefPRSNo.Text & "'")
    If Not rsPRS_Header.EOF And Not rsPRS_Header.BOF Then
        Set rsPRS_Details = gconDMIS.Execute("SELECT * FROM PMIS_VW_PRS_TRAN WHERE TRANNO = " & N2Str2Null(rsPRS_Header!TRANNO) & " AND STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text))
        If Not rsPRS_Details.EOF And Not rsPRS_Details.BOF Then
            SQL_STATEMENT = "UPDATE PMIS_VW_PRS_TRAN SET TREMARKS = 'SERVED'  WHERE TRANNO = " & N2Str2Null(rsPRS_Header!TRANNO) & " AND STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text)
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "EE", LOCALACESS, SQL_STATEMENT, labID, "ACCESSORIES", "TRAN NO: " & txtTranNo, COUNTERTYPE, labDetID
        End If
    End If
    rsRefresh
    On Error Resume Next
    RSORD_HD.Find "ID = " & labID.Caption
    StoreMemvars
    Screen.MousePointer = 0
    If ADDOREDIT = "ADD" Then
        cmdAddTran_Click
        fraDetails.Enabled = False
        Picture1.Enabled = False
    Else
        cmdTranCancel.Value = True
    End If
    Exit Sub
Errorcode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub



Private Sub cmd_add_ro_Click()
    PIC_ADD_RO.Visible = True
    PIC_ADD_RO.ZOrder 0
    Frame1.Enabled = False
    Text1_Change
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    If Module_Access(LOGID, "EDIT PARTS ISSUANCE AMOUNT", "SYSTEM") = False Then Exit Sub
    txtTranUPrice.Enabled = True
End Sub

Sub FillCboSalesMan()
    Set RSSALESMAN = New ADODB.Recordset
    RSSALESMAN.Open "select empno,signname from PMIS_vw_SalesMan order by signname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSSALESMAN.EOF And Not RSSALESMAN.BOF Then
        RSSALESMAN.MoveFirst: cboSalesMan.Clear: cboSMName.Clear
        Do While Not RSSALESMAN.EOF
            cboSalesMan.AddItem Null2String(RSSALESMAN!empno)
            cboSMName.AddItem Null2String(RSSALESMAN!signname)
            RSSALESMAN.MoveNext
        Loop
    Else
        cboSalesMan.Clear: cboSMName.Clear
    End If
End Sub

Sub FillDetails()
    KCNT = 0
    ORD_TOTUPRICE = 0
    ORD_TOTINVAMT = 0
    ORD_TOTVAT = 0
    ORD_TOTQTY = 0
    Dim STOCKDESCription                               As String
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "SELECT TRANTYPE,TRANNO,ID,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUPRICE FROM PMIS_TDAYTRAN WHERE TYPE='A' AND TRANNO = " & N2Str2Null(txtTranNo.Text) & " AND TRANTYPE = " & N2Str2Null(txtTranType.Text) & " ORDER BY ITEMNO ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        cboChargeTo.Enabled = False
        Screen.MousePointer = 11
        RSTDAYTRAN.MoveFirst
        Do While Not RSTDAYTRAN.EOF
            KCNT = KCNT + 1

            STOCKDESCription = SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_ORD))

            grdDetails.AddItem RSTDAYTRAN!ID & Chr(9) & Format(Null2String(RSTDAYTRAN!ITEMNO), "0000") & Chr(9) & _
                               Null2String(RSTDAYTRAN!STOCK_ORD) & Chr(9) & _
                               STOCKDESCription & Chr(9) & _
                               N2Str2IntZero(RSTDAYTRAN!TRANQTY) & Chr(9) & _
                               Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & Chr(9) & _
                               Format(N2Str2IntZero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT)
            ORD_TOTQTY = ORD_TOTQTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
            ORD_TOTUPRICE = ORD_TOTUPRICE + (N2Str2IntZero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
            ORD_TOTINVAMT = ORD_TOTINVAMT + (N2Str2IntZero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
            RSTDAYTRAN.MoveNext
        Loop
        If NumericVal(txtDS1.Text) <> 0 Then
            If txtDS_Desc1.Text = "" Then
                txtDS_Desc1.Text = "DISCOUNT"
            End If
            txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
            txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
        Else
            txtDS_Desc1.Text = ""
            txtDS_Amt1.Text = 0
            txtTTLInvAmt.Text = ToDoubleNumber(ORD_TOTUPRICE)
            txtNetInvAmt.Text = ToDoubleNumber(ORD_TOTINVAMT)
        End If
        ORD_TOTINVAMT = ORD_TOTINVAMT - NumericVal(txtDS_Amt1.Text)
        If KCNT <> 0 Then grdDetails.RemoveItem 1
        Screen.MousePointer = 0
    Else
        cboChargeTo.Enabled = True
        cleargrid grdDetails
    End If
End Sub


Function FillSalesMan(xxx As String) As String
    Set RSSALESMAN = New ADODB.Recordset
    RSSALESMAN.Open "select empno,signname from PMIS_vw_SalesMan where empno = '" & xxx & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSSALESMAN.EOF And Not RSSALESMAN.BOF Then
        FillSalesMan = Null2String(RSSALESMAN!signname)
        cboSalesMan.Text = Null2String(RSSALESMAN!empno)
    Else
        cboSalesMan.Text = ""
    End If
End Function


Sub X_FillSearchGrid(xxx As String)
    Dim RSORD_HD                                       As ADODB.Recordset
    Dim SQL_SEARCH_STRING                              As String
    lstOrd_Hd.Sorted = False
    lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False

    Set RSORD_HD = New ADODB.Recordset
    xxx = Replace(LTrim(RTrim(xxx)), "'", "")

    If optTranno.Value = True Then
        Set RSORD_HD = gconDMIS.Execute("SELECT TRANNO, ID FROM PMIS_ORD_HD WHERE TYPE='A' AND TRANTYPE = '" & COUNTERTYPE & "' AND TRANNO LIKE '" & xxx & "%' AND STATUS2='R'")
    ElseIf optRONo.Value = True Then
        Set RSORD_HD = gconDMIS.Execute("SELECT RONO, ID FROM PMIS_ORD_HD WHERE TYPE='A' AND TRANTYPE = '" & COUNTERTYPE & "' AND RONO LIKE '" & xxx & "%' AND STATUS2='R'  ORDER BY TRANNO ASC")
    Else
        Set RSORD_HD = gconDMIS.Execute("SELECT CUSTNAME, ID  FROM PMIS_ORD_HD WHERE TYPE='A' AND TRANTYPE = '" & COUNTERTYPE & "' AND CUSTNAME  LIKE '" & xxx & "%' AND STATUS2='R'  ORDER BY CUSTNAME")
    End If


    If Not (RSORD_HD.EOF And RSORD_HD.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HD
        lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub




Private Sub Command3_Click()
    If Module_Access(LOGID, "GENERATE NON INVOICE NUMBER", "DATA ENTRY") = False Then Exit Sub

    txtPRtranno.Visible = True
    txtPRtranno.SetFocus
    txtPRtranno.Locked = True
    Dim RSTMP                                          As New ADODB.Recordset
    Dim ISSCOUNTER                                     As Integer
    On Error GoTo Errorcode
    Set RSTMP = gconDMIS.Execute("SELECT COUNT(*) AS BILANG FROM PMIS_vw_ISS_HISTORY WHERE (TRANTYPE = 'RIV')  AND LEFT(TRANNO,1) = 'M' AND TYPE='A'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        ISSCOUNTER = NumericVal(RSTMP!BILANG)
    End If

    ISSCOUNTER = ISSCOUNTER + 1
    txtPRtranno.Text = "M" & Format(ISSCOUNTER, "00000")

    Set RSTMP = Nothing
Errorcode:
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim fild                                           As String
    Dim PCURONHAND                                     As Integer
    Dim PCurTISSQTY                                    As Integer
    Dim PCURISSUANCES                                  As Integer
    Dim rsTdaytranDup                                  As ADODB.Recordset
    Dim rsPartmasDup                                   As ADODB.Recordset
    Dim INVNUMBER                                      As String

    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    fild = grdDetails.Text

    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If picDetails.Visible = False Then Exit Sub

            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (Parts Issuance)"
            '====================================================================
            Call frmALL_AuditInquiry.DisplayHistory(labID, "PARTS ISSUANCE SERVICE ISSUANCE")
            '====================================================================

        Case vbKeyEscape
            If pic_viewStockADB.Visible = True Then
                pic_viewStockADB.Visible = False
            End If

            If Picture1.Visible = True Then
                SendToBack
                StoreMemvars
            End If
            If pic_viewStockADB.Visible = True Then
                pic_viewStockADB.Visible = False
            End If



            txtPRtranno.Visible = False
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(RSORD_HD!STATUS) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change..."
                ElseIf Null2String(RSORD_HD!STATUS) = "B" Then
                    MsgSpeechBox "Transactions are Already Billed-Out and cannot be Change..."
                ElseIf Null2String(RSORD_HD!STATUS) = "P" Then
                    MsgSpeechBox "Transactions are Already Posted and cannot be Change..."
                Else
                    cmdAddTran_Click
                    Picture1.Enabled = False
                    fraDetails.Enabled = False
                    picDetails.Enabled = False
                End If
            End If
        Case vbKeyF4
            If fild <> "" And fild <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(RSORD_HD!STATUS) <> "P" And Null2String(RSORD_HD!STATUS) <> "C" And Null2String(RSORD_HD!STATUS) <> "B" Then
                        grdDetails_DblClick
                        Picture1.Enabled = False
                        fraDetails.Enabled = False
                    End If
                End If
            End If
        Case vbKeyF5
            If fild <> "" And fild <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(RSORD_HD!STATUS) <> "P" And Null2String(RSORD_HD!STATUS) <> "C" And Null2String(RSORD_HD!STATUS) <> "B" Then
                        grdDetails_DblClick
                        cmdTranDelete_Click
                    End If
                End If
            End If
        Case vbKeyF8
            If cmdPost.Enabled = True Then
                cmdPost_Click
            End If
        Case vbKeyF11
            If Picture1.Visible = True And Null2String(RSORD_HD!STATUS) = "C" Then
                If MsgBox("Are you sure you want to uncancel this transaction", vbInformation + vbYesNo) = vbYes Then
                    gconDMIS.Execute ("UPDATE PMIS_ORD_HD SET STATUS=NULL WHERE ID=" & labID)
                    rsRefresh
                    RSORD_HD.Find ("id=" & labID)
                    StoreMemvars
                End If
            End If
        Case vbKeyF12

            If Picture1.Visible = True Then
                If Null2String(RSORD_HD!STATUS) = "B" Then
                    MsgBox "RO # " & txtRONO & " is already been invoiced. " & vbCrLf & "To Unpost this Transaction Service Invoice Should Be Cancelled First.", vbCritical
                    rsRefresh
                    RSORD_HD.Find ("ID=" & labID)
                    StoreMemvars
                    Exit Sub
                ElseIf Null2String(RSORD_HD!STATUS) = "P" Then
                    If Function_Access(LOGID, "Acess_UNPost", LOCALACESS) = False Then Exit Sub

                    If txtTranType = "RIV" Then
                        INVNUMBER = CheckIfROBilled(txtRONO)
                        If LTrim(RTrim(INVNUMBER)) <> "" Then
                            MsgBox "RO # " & txtRONO & " is already been invoiced. Service Invoice # " & INVNUMBER & vbCrLf & "Cannot Unpost Current Transaction." & vbCrLf & "To Unpost this Transaction Service Invoice Should Be Cancelled.", vbInformation
                            rsRefresh
                            RSORD_HD.Find ("ID=" & labID)
                            StoreMemvars
                            'Exit Sub
                        End If
                    End If

                    Set rsTdaytranDup = New ADODB.Recordset
                    rsTdaytranDup.Open "SELECT ID,TRANTYPE,TRANNO,STOCK_ORD,TRANQTY FROM PMIS_TDAYTRAN WHERE TYPE='A' AND TRANNO = " & N2Str2Null(RSORD_HD!TRANNO) & " AND TRANTYPE = " & N2Str2Null(RSORD_HD!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
                    If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
                        rsTdaytranDup.MoveFirst
                        Do While Not rsTdaytranDup.EOF
                            Set rsPartmasDup = New ADODB.Recordset
                            rsPartmasDup.Open "SELECT STOCKNO,ONHAND,TISSQTY,TISSQTY,ISSUANCES,REQSERVED,S_REQSERVED FROM PMIS_STOCKMAS WHERE TYPE='M' AND  STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD), gconDMIS
                            If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                                If COUNTERTYPE <> "ADB" Then
                                    PCURONHAND = N2Str2IntZero(rsPartmasDup!ONHAND) + N2Str2Zero(rsTdaytranDup!TRANQTY)
                                    PCurTISSQTY = N2Str2IntZero(rsPartmasDup!TISSQTY) - N2Str2Zero(rsTdaytranDup!TRANQTY)
                                    PCURISSUANCES = N2Str2IntZero(rsPartmasDup!ISSUANCES) - N2Str2Zero(rsTdaytranDup!TRANQTY)

                                    If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                                        SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                                      " REQSERVED = " & N2Str2IntZero(rsPartmasDup!REQSERVED) - N2Str2Zero(rsTdaytranDup!TRANQTY) & _
                                                      " WHERE TYPE='M' AND STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                                        gconDMIS.Execute SQL_STATEMENT
                                        Call NEW_LogAudit("E", "ACCESSORIES MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "A", "TRAN NO : " & Null2String(RSORD_HD!TranType) & " - UNPOST", COUNTERTYPE, "")
                                    Else
                                        SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                                      " S_REQSERVED = " & N2Str2IntZero(rsPartmasDup!S_REQSERVED) - N2Str2Zero(rsTdaytranDup!TRANQTY) & _
                                                      " WHERE TYPE='M' AND STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                                        gconDMIS.Execute SQL_STATEMENT
                                        Call NEW_LogAudit("E", "ACCESSORIES MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "A", "TRAN NO : " & Null2String(RSORD_HD!TranType & " - UNPOST"), COUNTERTYPE, "")
                                    End If


                                    SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                                  " ONHAND = " & PCURONHAND & "," & _
                                                  " TISSQTY = " & PCurTISSQTY & "," & _
                                                  " ISSUANCES = " & PCURISSUANCES & "," & _
                                                  " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                                                  " LASTUPDATE = '" & LOGDATE & "'" & _
                                                  " WHERE TYPE='M' AND STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                                    gconDMIS.Execute SQL_STATEMENT
                                    Call NEW_LogAudit("E", "ACCESSORIES MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "A", "TRAN NO : " & Null2String(RSORD_HD!TranType) & " - UNPOST", COUNTERTYPE, "")
                                End If

                                SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET" & _
                                              " STATUS = 'N'," & _
                                              " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                                              " LASTUPDATE = '" & LOGDATE & "'" & _
                                              " WHERE ID = " & rsTdaytranDup!ID
                                gconDMIS.Execute SQL_STATEMENT
                                NEW_LogAudit "UU", LOCALACESS, SQL_STATEMENT, labID, "ACCESSORIES", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""
                            End If
                            rsTdaytranDup.MoveNext
                        Loop
                    End If
                    SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                                  " STATUS = 'N'," & _
                                  " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                                  " LASTUPDATE = '" & LOGDATE & "'" & _
                                  " WHERE ID = " & labID.Caption

                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "U", LOCALACESS, SQL_STATEMENT, labID, "ACCESSORIES", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""

                    gconDMIS.Execute ("UPDATE PMIS_ORD_HD SET STATUS3=NULL WHERE RONO=" & N2Str2Null(txtRONO))
                    gconDMIS.Execute ("UPDATE PMIS_ORD_HIST SET STATUS3=NULL WHERE RONO=" & N2Str2Null(txtRONO))

                    rsRefresh
                    RSORD_HD.Find "id =" & labID.Caption
                    StoreMemvars
                End If
                Set rsTdaytranDup = Nothing
                Set rsPartmasDup = Nothing
            End If

        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    LOCALACESS = "ACCESSORIES SERVICE ISSUANCE"
    CenterMe frmMain, Me, 1
    PMIS_ORDER_SHOW = True
    textSearch.Text = ""

    If COUNTERTYPE = "CSH" Then optCASH.Value = True
    If COUNTERTYPE = "CHG" Then optCHARGE.Value = True

    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False



    InitMemVars
    txtTranUPrice.Enabled = False
    rsRefresh
    On Error Resume Next
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        RSORD_HD.MoveLast
    End If
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PMIS_ORDER_SHOW = False
    COUNTERTYPE = ""
    Unload Me
End Sub

Private Sub grdDetails_DblClick()
    Dim fild                                           As String
    If Null2String(RSORD_HD!STATUS) = "C" Then
        MsgSpeech "Transactions are Already Cancelled and cannot be Change"

        MsgBox "Transactions are Already Cancelled" & vbCrLf & _
               "and cannot be Change", vbInformation, "Edit Not Allowed!"

    ElseIf Null2String(RSORD_HD!STATUS) = "B" Then
        MsgSpeech "Transactions are Already Billed-Out and cannot be Change"

        MsgBox "Transactions are Already Billed-Out" & vbCrLf & _
               "and Cannot be Changed", vbInformation, "Edit Not Allowed!"
    ElseIf Null2String(RSORD_HD!STATUS) = "P" Then
        MsgSpeech "Transactions are Already Posted and cannot be Change"
        MsgBox "Transactions are Already Posted" & vbCrLf & _
               "and Cannot Be Changed!", vbInformation, "Edit Not Allowed!"
    Else
        grdDetails.Row = grdDetails.Row
        grdDetails.Col = 0
        fild = grdDetails.Text
        If fild <> "" And fild <> "No Entry" Then
            ADDOREDIT = "EDIT"
            cmdTranDelete.Enabled = True
            BringToFront
            StorePartsEntry (fild)
        Else
            MsgSpeechBox "No Entry on Parts!"
            Exit Sub
        End If
    End If
End Sub

Private Sub grdDetails_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Sub InitCbo()
    Set RSPARTMAS = New ADODB.Recordset

    RSPARTMAS.Open "SELECT DISTINCT STOCK_ORD FROM PMIS_ALLDAYTRAN INNER JOIN PMIS_STOCKMAS " & _
                 " ON PMIS_STOCKMAS.TYPE=PMIS_ALLDAYTRAN.TYPE AND PMIS_ALLDAYTRAN.STOCK_ORD=PMIS_STOCKMAS.STOCKNO WHERE PMIS_STOCKMAS.TYPE='M' AND " & _
                 " (TRANNO  IN (SELECT TRANNO FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='M' AND RONO='" & txtRONO & "' AND ISNULL(STATUS2,'')  <>'R' AND (STATUS='P' OR STATUS='B')) " & _
                 " OR TRANNO IN (SELECT TRANNO FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='M' AND RONO='" & txtRONO & "' AND ISNULL(STATUS2,'')  <>'R' AND (STATUS='P' OR STATUS='B')))   " & _
                 " AND TRANTYPE='ADB' AND ONHAND>0", gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        RSPARTMAS.MoveFirst
        cboTranPartNo.Clear
        Do While Not RSPARTMAS.EOF
            cboTranPartNo.AddItem Null2String(RSPARTMAS!STOCK_ORD)
            RSPARTMAS.MoveNext
        Loop
    End If
    FillCboSalesMan
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

Sub InitGrid()
    With grdDetails
        .Rows = 7
        .ColWidth(0) = 1
        .ColWidth(1) = 1000
        .ColWidth(2) = 1500
        .ColAlignment(2) = 2
        .ColWidth(3) = 2200
        .ColWidth(4) = 1000
        .ColWidth(5) = 1200
        .ColWidth(6) = 1300
        .Row = 0
        .Col = 1
        .Text = "Item"
        .Col = 2
        .Text = "Stock Number"
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

Sub InitMemVars()


    Set RSCUNTER = New ADODB.Recordset
    RSCUNTER.Open "select * from PMIS_Counter where TYPE='A' AND modul = 'RIV'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSCUNTER.EOF And Not RSCUNTER.BOF Then
        txtTranNo.Text = Format(Null2String(RSCUNTER!nextnumber), "000000")
    Else
        txtTranNo.Text = "000001"
    End If
    txtRONO.Enabled = True
    txtTerms.Enabled = False


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
    InitGrid
    txtTranDate.Enabled = False
    cleargrid grdDetails
    SendToBack
    InitSignatories
End Sub

Sub InitParts()
    txtTranItemNo.Text = Format(KCNT + 1, "0000")
    cboTranPartNo.Text = ""
    txtTranDescription.Text = ""
    txtTranQty.Text = 1
    txtTranUCost.Text = "0.00"
    txtTranUPrice.Text = "0.00"
    txtTranTotalAmt.Text = "0.00"
    labTranUCost.Visible = False
    txtTranUCost.Visible = False
    Check1.Enabled = False
End Sub

Sub InitSignatories()
    txtPreparedBy.Text = ""
    txtIssuedBy.Text = ""
    txtRequestedBy.Text = ""
    txtApprovedBy.Text = ""
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



Private Sub lstOrd_Hd_GotFocus()
    On Error Resume Next
    lstOrd_Hd_ItemClick lstOrd_Hd.SelectedItem

End Sub

Private Sub lstOrd_Hd_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Clipboard.SetText lstOrd_Hd.SelectedItem.Text
    RSORD_HD.MoveFirst
    RSORD_HD.Find ("ID=" & Item.ListSubItems(1).Text)
    StoreMemvars
End Sub

Private Sub lstOrd_Hd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub


Private Sub Option1_Click()
    lstOrd_Hd.ColumnHeaders(1).Text = "CUSTOMER NAME"
    X_FillSearchGrid ""
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optRONo_Click()
    lstOrd_Hd.ColumnHeaders(1).Text = "RO Number"
    X_FillSearchGrid ""
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optTranno_Click()
    lstOrd_Hd.ColumnHeaders(1).Text = "Tran. No."
    X_FillSearchGrid ""
    On Error Resume Next
    textSearch.SetFocus
End Sub

Sub PISPRINTING()
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

Sub rsRefresh()
    Set RSORD_HD = New ADODB.Recordset
    RSORD_HD.Open "SELECT * FROM PMIS_ORD_HD WHERE TYPE='A' AND TRANTYPE = 'RIV' AND STATUS2='R' ORDER BY TRANNO ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    InitCboChargeToCounter
End Sub

Sub SendToBack()
    fraAddTran.ZOrder 1
    fraAddTran.Visible = False
    fraAddTran.Enabled = False
    fraSignatories.ZOrder 1
    fraSignatories.Visible = False
    Picture1.Enabled = True
    fraDetails.Enabled = True
    picDetails.Enabled = True
End Sub

Sub SERVICEPISPRINTING()
    Screen.MousePointer = 11
    If NumericVal(txtDS1.Text) = 0 Then
        rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
        rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

        If RSORD_HD!TranType = "RIV" Then
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Parts.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'RIV' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Parts.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'ADB' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        End If
    Else
        If RSORD_HD!TranType = "RIV" Then
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Partsdisc.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'RIV' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Partsdisc.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'ADB' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        End If
    End If


    Screen.MousePointer = 0

End Sub

Sub SERVICEPISPRINTING_BLANKFORM()
    Screen.MousePointer = 11
    Dim cnt1, cnt2, cnt3                               As Integer
    Dim knt, cntCOPY                                   As Integer
    Dim TOTALQTY, TOTALPRICE                           As Double
    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile where ModuleName = 'PMIS'", gconDMIS

    Open App.Path & "\PIS.HTML" For Output As #1
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where TYPE = 'P' AND tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = 'RIV' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        TOTALQTY = 0
        TOTALPRICE = 0

        If COMPANY_CODE = "HAI" Then
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then cntCOPY = 4 Else cntCOPY = 2
        Else
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then cntCOPY = 4 Else cntCOPY = 1
        End If


        Print #1, "<html><body>"
        knt = 0
        For knt = 1 To cntCOPY
            If knt < 3 Then
                RSTDAYTRAN.MoveFirst
                TOTALQTY = 0: TOTALPRICE = 0
            Else
                If RSTDAYTRAN.EOF Then
                    RSTDAYTRAN.MoveLast
                Else
                    RSTDAYTRAN.MoveNext
                End If
            End If
            Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
            Print #1, "<td align=center width=60%><font size=3 FACE=TIMES NEW ROMAN>" & RSPROFILE!CompanyName & "</font></td>"
            Print #1, "<td align=right width=20%><font size=1 FACE=TIMES NEW ROMAN>COPY: " & knt & "</font></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>PARTS ISSUANCE SLIP</strong></font></td>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "<td align=center width=60%>&nbsp;</td>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"

            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Repair Order Number:&nbsp;</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & Null2String(RSORD_HD!RoNo) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%>&nbsp;</td>"
            Print #1, "</tr>"

            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & "SERVICE PIS-" & Null2String(RSORD_HD!TRANNO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(RSORD_HD!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(RSORD_HD!CUSTCODE) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(RSORD_HD!CUSTNAME) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=5%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ITM #</FONT></td>"
            Print #1, "<td width=20%><FONT SIZE=2 FACE=TIMES NEW ROMAN>PART NUMBER</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>DESCRIPTION</FONT></td>"
            Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>QTY</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>UNIT PRICE</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>TOTAL PRICE</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            cnt1 = 0
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then
                cnt2 = 0
            Else
                cnt2 = MAX_ISS_LINE - RSTDAYTRAN.RecordCount
            End If
            If knt >= 3 Then cnt2 = MAX_ISS_LINE - (RSTDAYTRAN.RecordCount - MAX_ISS_LINE)
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            If RSTDAYTRAN.AbsolutePosition > MAX_ISS_LINE Then
                RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE + 1
            End If
            Do While Not RSTDAYTRAN.EOF
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!ITEMNO) & "</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!STOCK_ORD) & "</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_SUP)) & "</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(RSTDAYTRAN!TRANQTY) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE)
                End If
                Print #1, "</tr>"
                If RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE Then Exit Do
                RSTDAYTRAN.MoveNext
            Loop
            For cnt3 = 1 To cnt2
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "</tr>"
            Next
            Print #1, "</table>"
            '==================================
            'updating code:     JAA  - 02092008
            If COMPANY_CODE = "HBK" Then
                Print #1, "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
            End If
            '==================================
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            If cntCOPY = 4 And knt < 3 Then
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            Else
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL PIS</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & TOTALQTY & "</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & Format(TOTALPRICE, MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtPreparedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtIssuedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtApprovedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtRequestedBy.Text & "</FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Requested By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Approved By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Issued By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Received By</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            If knt <> 2 And knt <> 4 Then
                Print #1, "<table>"
                Print #1, "<tr>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
                'Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
                Print #1, "<table>"
                Print #1, "<tr>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
        Next
        Print #1, "</body></html>"
        Close #1
        On Error Resume Next
        Open App.Path & "\PIS.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRIV.Navigate "about:blank"

            browRIV.Refresh
            browRIV.Navigate App.Path & "\PIS.HTML"
            DoEvents
            browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            Screen.MousePointer = 0
        End If
    End If
    Set RSPROFILE = Nothing
    Screen.MousePointer = 0
End Sub

Sub SetCustInfo(rep As String)
    Dim RSCUSTINFO                                     As ADODB.Recordset
    rep = Left(rep, 1) & "-" & Right(rep, 6)
    Set RSREPOR = New ADODB.Recordset
    RSREPOR.Open "SELECT REP_OR,NIYM,ACCT_NO,INVOICE,PLATE_NO,DTE_REL FROM CSMS_REPOR WHERE REP_OR = '" & txtRONO.Text & "'", gconDMIS


    If Not RSREPOR.EOF And Not RSREPOR.BOF Then
        If Null2String(RSREPOR!Invoice) <> "" Then
            REPOR_STATUS = "Billed-Out"
        End If
        txtCustName.Text = Null2String(RSREPOR!niym)
        txtCustCode.Text = Null2String(RSREPOR!ACCT_NO)

        If Null2String(RSREPOR!plate_no) <> "" Then
            Set RSCUSTINFO = New ADODB.Recordset
            'IT MIGHT GIVE A WRONG INFO OF THE CUSTOMER
            Set RSCUSTINFO = gconDMIS.Execute("SELECT * FROM CSMS_CUSVEH WHERE PLATE_NO=" & N2Str2Null(RSREPOR!plate_no) & "  AND CUSCDE='" & Null2String(RSREPOR!ACCT_NO) & "'")
            If Not RSCUSTINFO.EOF Or Not RSCUSTINFO.BOF Then
                txtRemarks = "MODEL: " & Null2String(RSCUSTINFO("MODEL")) & vbCrLf & "ENGINE#:" & Null2String(RSCUSTINFO("SERIAL")) & vbCrLf & "VIN#:" & Null2String(RSCUSTINFO("VIN")) & vbCrLf & "PLATE#:" & Null2String(RSCUSTINFO("PLATE_NO"))
            End If
        End If
    Else
        txtCustName.Text = ""
        txtCustCode.Text = ""

    End If
End Sub

Sub SetCustomer()
    Dim RSCUSTOMER                                     As ADODB.Recordset
    Set RSCUSTOMER = New ADODB.Recordset
    Set RSCUSTOMER = gconDMIS.Execute("SELECT * FROM ALL_CUSTOMER WHERE CUSCDE = '" & txtCustCode.Text & "'")
    If Not RSCUSTOMER.EOF And Not RSCUSTOMER.BOF Then
        txtCustName.Text = Null2String(RSCUSTOMER!AcctName) & vbCrLf & Null2String(RSCUSTOMER!CUSTOMERADD) & vbCrLf & Null2String(RSCUSTOMER!City)
    Else
        txtCustName = ""
    End If
End Sub

Sub SetPartDetails(xxx As String)
    Dim RSPARTMAS                                      As ADODB.Recordset
    Dim STR_SQLX                                       As String
    Dim RSADBISSUE                                     As ADODB.Recordset
    Dim RSADB_FILL                                     As ADODB.Recordset

    Set RSPARTMAS = gconDMIS.Execute("SELECT * FROM PMIS_STOCKMAS WHERE TYPE='A' AND STOCKNO = '" & xxx & "' AND ACTIVE = 'Y'")
    TXT_ADB_ISSUE = 0
    TXT_ADB_FILL = 0
    TXT_AVL_ISSUE = 0
    TXT_CURRONHAND = 0

    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        If N2Str2Zero(RSPARTMAS!ONHAND) > 0 Then chkAvailableOnStock.Value = 1 Else chkAvailableOnStock.Value = 0

        TXT_CURRONHAND = N2Str2Zero(RSPARTMAS!ONHAND)

        STR_SQLX = "SELECT 'CURRENT', SUM(TRANQTY) AS TQTY FROM PMIS_TDAYTRAN WHERE TRANNO  IN"
        STR_SQLX = STR_SQLX & "(SELECT TRANNO FROM PMIS_ORD_HD WHERE TYPE='M' AND  TRANTYPE='ADB' AND RONO='" & txtRONO & "' AND ISNULL(STATUS2,'')  <>'R' AND (STATUS='P' OR STATUS='B'))"
        STR_SQLX = STR_SQLX & "AND TRANTYPE='ADB'"
        STR_SQLX = STR_SQLX & "AND STOCK_ORD='" & xxx & "'"
        STR_SQLX = STR_SQLX & "GROUP BY STOCK_ORD"
        STR_SQLX = STR_SQLX & " Union " & vbCrLf
        STR_SQLX = STR_SQLX & "SELECT 'OLD', SUM(TRANQTY) AS TQTY FROM PMIS_DAYTRAN WHERE TRANNO  IN"
        STR_SQLX = STR_SQLX & "(SELECT TRANNO FROM PMIS_ORD_HIST WHERE TYPE='M' AND TRANTYPE='ADB' AND RONO='" & txtRONO & "' AND ISNULL(STATUS2,'')  <>'R' AND (STATUS='P' OR STATUS='B'))"
        STR_SQLX = STR_SQLX & "AND TRANTYPE='ADB'"
        STR_SQLX = STR_SQLX & "AND STOCK_ORD='" & xxx & "'"
        STR_SQLX = STR_SQLX & "GROUP BY STOCK_ORD"
        Set RSADBISSUE = gconDMIS.Execute(STR_SQLX)

        While Not RSADBISSUE.EOF
            TXT_ADB_ISSUE = N2Str2Zero(TXT_ADB_ISSUE) + N2Str2Zero(RSADBISSUE!TQTY)
            RSADBISSUE.MoveNext
        Wend

        STR_SQLX = "SELECT 'CURRENT', SUM(TRANQTY) AS TQTY FROM PMIS_TDAYTRAN WHERE TRANNO  IN"
        STR_SQLX = STR_SQLX & "(SELECT TRANNO FROM PMIS_ORD_HD WHERE TYPE='M' AND  TRANTYPE='RIV' AND RONO='" & txtRONO & "' AND ISNULL(STATUS2,'')  ='R' AND (STATUS='P' OR STATUS='B'))"
        STR_SQLX = STR_SQLX & "AND TRANTYPE='RIV'"
        STR_SQLX = STR_SQLX & "AND STOCK_ORD='" & xxx & "'"
        STR_SQLX = STR_SQLX & "GROUP BY STOCK_ORD"
        STR_SQLX = STR_SQLX & " Union " & vbCrLf
        STR_SQLX = STR_SQLX & "SELECT 'OLD', SUM(TRANQTY) AS TQTY FROM PMIS_DAYTRAN WHERE TRANNO  IN"
        STR_SQLX = STR_SQLX & "(SELECT TRANNO FROM PMIS_ORD_HIST WHERE TYPE='M' AND TRANTYPE='RIV' AND RONO='" & txtRONO & "' AND ISNULL(STATUS2,'')  ='R' AND (STATUS='P' OR STATUS='B'))"
        STR_SQLX = STR_SQLX & "AND TRANTYPE='RIV'"
        STR_SQLX = STR_SQLX & "AND STOCK_ORD='" & xxx & "'"
        STR_SQLX = STR_SQLX & "GROUP BY STOCK_ORD"

        Set RSADB_FILL = gconDMIS.Execute(STR_SQLX)

        While Not RSADB_FILL.EOF
            TXT_ADB_FILL = N2Str2Zero(TXT_ADB_FILL) + N2Str2Zero(RSADB_FILL!TQTY)
            RSADB_FILL.MoveNext
        Wend

        TXT_AVL_ISSUE = 0

        If TXT_CURRONHAND <= N2Str2Zero(TXT_ADB_ISSUE) - N2Str2Zero(TXT_ADB_FILL) Then
            TXT_AVL_ISSUE = TXT_ADB_ISSUE
        Else
            TXT_AVL_ISSUE = N2Str2Zero(TXT_ADB_ISSUE) - N2Str2Zero(TXT_ADB_FILL)
        End If
    End If
End Sub

Function SetPartIDDesc(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "SELECT ID,STOCKDESC FROM PMIS_STOCKMAS WHERE TYPE='M' AND LTRIM(RTRIM(STOCKDESC)) = '" & LTrim(RTrim(DDD)) & "' AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDDesc = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartIDSTOCKNO(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "SELECT ID,STOCKNO FROM PMIS_STOCKMAS WHERE TYPE='M' AND STOCKNO = " & N2Str2Null(DDD) & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDSTOCKNO = Null2String(RSPARTMAS!ID)
        SetPartDetails DDD
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "SELECT SRP,STOCKNO,MAC,DNP FROM PMIS_STOCKMAS WHERE TYPE='M' AND STOCKNO = '" & ppp & "' AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then

            If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                SetPartPrice = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac) * 1.12)
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            Else
                SetPartPrice = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            End If

        End If
        SetPartDetails ppp
    End If
End Function

Function SetStockCost(xxx As String) As Double
    Dim rsSTKCost                                      As ADODB.Recordset
    Set rsSTKCost = New ADODB.Recordset
    Set rsSTKCost = gconDMIS.Execute("SELECT MAC FROM PMIS_STOCKMAS WHERE TYPE='M' AND STOCKNO = '" & xxx & "'")
    If Not rsSTKCost.EOF And Not rsSTKCost.BOF Then
        SetStockCost = N2Str2Zero(rsSTKCost!Mac)
    End If
End Function

Function SetSTOCKDESC(ppp As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "SELECT STOCKNO,STOCKDESC,SRP,MAC,DNP FROM PMIS_STOCKMAS WHERE TYPE='M' AND STOCKNO= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKDESC = Null2String(RSPARTMAS!STOCKDESC)
        If txtTranType.Text = "DR" Then
            If cboChargeTo.Text = "PARTS CLAIM" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            End If
        Else
            If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!dnp))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            ElseIf Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            End If




        End If
    Else
        txtTranUPrice.Text = 0
        txtTranUCost.Text = 0
    End If
End Function

Function SetSTOCKDESC2(pid As Variant)
    Dim rsPRS_Header                                   As ADODB.Recordset
    Dim rsPRS_Details                                  As ADODB.Recordset

    If pid <> "" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "SELECT ID,STOCKDESC,SRP,MAC,DNP FROM PMIS_STOCKMAS WHERE ID = " & pid & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            SetSTOCKDESC2 = Null2String(RSPARTMAS!STOCKDESC)
            If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!dnp))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            ElseIf Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            Else
                Set rsPRS_Header = gconDMIS.Execute("SELECT * FROM PMIS_VW_PRS WHERE REFPISNO = '" & cboRefPRSNo.Text & "'")
                If Not rsPRS_Header.EOF And Not rsPRS_Header.BOF Then
                    Set rsPRS_Details = gconDMIS.Execute("SELECT * FROM PMIS_VW_PRS_TRAN WHERE TRANNO = " & N2Str2Null(rsPRS_Header!TRANNO) & " AND STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text))
                    If Not rsPRS_Details.EOF And Not rsPRS_Details.BOF Then
                        txtTranQty.Text = N2Str2Zero(rsPRS_Details!TRANQTY)
                    End If
                End If
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            End If
        Else
            txtTranUPrice.Text = "0.00"
            txtTranUCost.Text = 0
        End If
    End If

End Function

Function SetSTOCKNO(pid As Variant)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "SELECT ID,STOCKNO,SRP,DNP,MAC FROM PMIS_STOCKMAS WHERE ID = " & pid & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKNO = Null2String(RSPARTMAS!STOCKNO)
        If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
            txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!dnp))
            txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
        Else
            txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
            txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
        End If
    Else
        txtTranUPrice.Text = "0.00"
        txtTranUCost.Text = 0
    End If
End Function

Sub StoreMemvars()
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        labID.Caption = RSORD_HD!ID
        txtTranType.Text = Null2String(RSORD_HD!TranType)
        cboSMName.Enabled = True
        txtTranNo.Text = Null2String(RSORD_HD!TRANNO)
        txtTranDate.Text = Null2String(RSORD_HD!trandate)
        txtCustCode.Text = Null2String(RSORD_HD!CUSTCODE)
        txtCustName.Text = Null2String(RSORD_HD!CUSTNAME)
        txtReferencePIS.Text = Null2String(RSORD_HD!REFPISNO)
        cboRefPRSNo.Text = Null2String(RSORD_HD!refpRsno)

        If Mid(txtReferencePIS, 5, 1) = "W" Then
            txtTranUPrice.Enabled = True
        Else
            txtTranUPrice.Enabled = False
        End If

        If Null2String(RSORD_HD!chargeto) = "MEC" Then
            cboChargeTo.Text = "MECHANICAL"
        ElseIf Null2String(RSORD_HD!chargeto) = "COM" Then
            cboChargeTo.Text = "COMPANY"
        ElseIf Null2String(RSORD_HD!chargeto) = "WAR" Then
            cboChargeTo.Text = "WARRANTY"
        ElseIf Null2String(RSORD_HD!chargeto) = "TIN" Then
            cboChargeTo.Text = "TINSMITH"
        ElseIf Null2String(RSORD_HD!chargeto) = "FLE" Then
            cboChargeTo.Text = "FLEET"
        ElseIf Null2String(RSORD_HD!chargeto) = "VAR" Then
            cboChargeTo.Text = "VARIOUS"
        ElseIf Null2String(RSORD_HD!chargeto) = "PCL" Then
            cboChargeTo.Text = "PARTS CLAIM"
        Else
            cboChargeTo.Text = ""
        End If
        txtRONO.Text = Null2String(RSORD_HD!RoNo)
        cboSMName.Text = FillSalesMan(Null2String(RSORD_HD!salesman))
        txtTerms.Text = Null2String(RSORD_HD!TERMS)
        txtTTLInvAmt.Text = ToDoubleNumber(N2Str2Zero(RSORD_HD!ttlinvamt))
        txtDS1.Text = N2Str2IntZero(RSORD_HD!ds1)
        txtDS_Desc1.Text = Null2String(RSORD_HD!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(RSORD_HD!ds_amt1))
        txtNetInvAmt.Text = ToDoubleNumber(N2Str2Zero(RSORD_HD!netinvamt))
        txtRemarks.Text = Null2String(RSORD_HD!REMARKS)

        If Null2String(RSORD_HD!STATUS2) = "R" Then
            LAB_ADB = "ISSUANCE AGAINST ADB"
        Else
            LAB_ADB = ""
        End If

        If Null2String(RSORD_HD!STATUS) = "C" Then
            labPosted.Caption = "CANCELLED"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = False
        ElseIf Null2String(RSORD_HD!STATUS) = "P" Then
            labPosted.Caption = "POSTED"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = True
        Else
            labPosted.Caption = ""
            cmdEdit.Enabled = True
            cmdCancelCO.Enabled = True
            cmdPost.Enabled = True
            cmdPrint.Enabled = False
        End If



        cleargrid grdDetails
        FillDetails
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Function StorePartsEntry(ByVal ID As Variant)
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "SELECT ID,STOCK_ORD,STOCK_SUP,TRANQTY,ITEMNO,TRANUPRICE,TRANUCOST FROM PMIS_TDAYTRAN WHERE ID = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        labDetID.Caption = RSTDAYTRAN!ID
        labPartNo.Caption = Null2String(RSTDAYTRAN!STOCK_ORD)
        labPrevOrdQty.Caption = N2Str2IntZero(RSTDAYTRAN!TRANQTY)
        txtTranItemNo.Text = Format(Null2String(RSTDAYTRAN!ITEMNO), "0000")
        cboTranPartNo.Text = Null2String(RSTDAYTRAN!STOCK_ORD)
        txtTranDescription.Text = SetSTOCKDESC(RSTDAYTRAN!STOCK_ORD)
        txtTranQty.Text = N2Str2IntZero(RSTDAYTRAN!TRANQTY)
        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
        txtTranTotalAmt.Text = ToDoubleNumber(N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
        txtTranUPrice.Enabled = False
    End If

    labTranUCost.Visible = False
    txtTranUCost.Visible = False

End Function

Private Sub textSearch_Change()
    If optTranno.Value = True Then
        X_FillSearchGrid textSearch.Text
    ElseIf optRONo.Value = True Then
        Dim RONOStr                                    As String
        RONOStr = textSearch.Text
        If Left(RONOStr, 2) = "R-" Then
            RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
        Else
            RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
        End If
        X_FillSearchGrid RONOStr
    Else
        X_FillSearchGrid textSearch
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstOrd_Hd.ListItems.Count > 0 And lstOrd_Hd.Enabled = True Then: lstOrd_Hd.SetFocus
    End If
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

Private Sub txtCustName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
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

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtTranDate_LostFocus()
    txtTranDate.Text = Format(txtTranDate.Text, "SHORT DATE")
    If IsDate(txtTranDate) = True Then
        If DateDiff("m", txtTranDate, LOGDATE) <> 0 Then
            MsgBox "Warning: Transaction Month cannot be greater or less than the current month.", vbCritical
            txtTranDate.SetFocus
        End If
    Else
        MsgBox "Warning: Please Input Valid Date.", vbCritical
    End If
End Sub

Private Sub txtTranNo_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtTranNo_LostFocus()
    txtTranNo.Text = Format(txtTranNo.Text, "000000")
    Dim RSFINDDUP                                      As ADODB.Recordset
    If ADDOREDIT = "ADD" Then
        Set RSFINDDUP = New ADODB.Recordset
        RSFINDDUP.Open "select trantype,tranno from PMIS_Ord_Hd where TYPE='A' AND trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
            MsgSpeechBox "Transaction No. already exist!"
            On Error Resume Next
            Exit Sub
        End If
    Else
        If LTrim(RTrim(txtTranNo)) <> LTrim(RTrim(Null2String(RSORD_HD!TRANNO))) Then
            Set RSFINDDUP = New ADODB.Recordset
            RSFINDDUP.Open "select trantype,tranno from PMIS_Ord_Hd where TYPE='A' AND trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                MsgSpeechBox "Transaction No. already exist!"

                Exit Sub
            End If
        End If
    End If

End Sub

Private Sub txttranQty_Change()
    If txtTranQty.Text <> "" Then
        'EAP:012709 Validation for negative and zero issuances.
        If txtTranQty.Text <= 0 Then
            MessagePop InfoVoid, "Invalid Input", "Quantity must not have a zero or negative value"

            On Error Resume Next
            txtTranQty.SetFocus
            cmdTranSave.Enabled = False
            Exit Sub
        Else
            cmdTranSave.Enabled = True
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
        End If

    End If


End Sub

Private Sub txtTranQty_GotFocus()
    If NumericVal(txtTranQty.Text) = 1 Then txtTranQty.Text = ""
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


Private Sub txtTranTotalAmt_Change()
    txtTranTotalAmt.Text = Format(txtTranTotalAmt.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtTranUPrice_Change()
    If txtTranUPrice.Text <> "" Then
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
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


Public Sub PisValidation()
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    txtTranDate.Enabled = False
    cmdPost.Enabled = False
    'EAP:033109 so user cannot pressd f8 when transaction is not yet saved.
    cmdPost.Enabled = False
End Sub
