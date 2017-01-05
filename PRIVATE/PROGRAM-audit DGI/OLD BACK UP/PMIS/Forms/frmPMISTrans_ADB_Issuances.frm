VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmPMISTrans_ADB_Issuances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issuances Against Advance Bill"
   ClientHeight    =   7155
   ClientLeft      =   1110
   ClientTop       =   2520
   ClientWidth     =   11460
   ForeColor       =   &H00DEDFDE&
   Icon            =   "frmPMISTrans_ADB_Issuances.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11460
   Begin VB.PictureBox fraAddTran 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3555
      Left            =   1890
      ScaleHeight     =   3525
      ScaleWidth      =   9375
      TabIndex        =   41
      Top             =   1920
      Width           =   9405
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
         Left            =   6780
         TabIndex        =   126
         Top             =   2280
         Width           =   2205
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
         Left            =   7620
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   125
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   600
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
         Left            =   7680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   123
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   1710
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
         Left            =   7620
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   120
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   1020
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
         Left            =   7620
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   118
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   3150
         TabIndex        =   108
         Top             =   1890
         Width           =   285
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
         Height          =   3345
         Left            =   3840
         TabIndex        =   91
         Top             =   0
         Width           =   2865
         Begin VB.Frame Frame5 
            Caption         =   "Model Codes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   150
            TabIndex        =   106
            Top             =   2400
            Width           =   2595
            Begin VB.TextBox txtModelCode 
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
               Height          =   375
               Left            =   120
               MaxLength       =   6
               TabIndex        =   107
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
            Height          =   465
            Left            =   120
            TabIndex        =   105
            Top             =   180
            Width           =   2685
         End
         Begin VB.Frame Frame3 
            Height          =   975
            Left            =   150
            TabIndex        =   92
            Top             =   630
            Width           =   2595
            Begin VB.OptionButton optConsigned 
               Caption         =   "Consigned"
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
               Height          =   255
               Left            =   150
               TabIndex        =   95
               Top             =   660
               Width           =   1845
            End
            Begin VB.OptionButton optImported 
               Caption         =   "Imported"
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
               Height          =   255
               Left            =   150
               TabIndex        =   94
               Top             =   390
               Value           =   -1  'True
               Width           =   1845
            End
            Begin VB.OptionButton optLocalPurchase 
               Caption         =   "Local Purchases"
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
               Height          =   255
               Left            =   150
               TabIndex        =   93
               Top             =   150
               Width           =   1845
            End
         End
         Begin VB.Frame Frame4 
            Height          =   765
            Left            =   150
            TabIndex        =   96
            Top             =   1590
            Width           =   2595
            Begin VB.OptionButton optGenuine 
               Caption         =   "Genuine"
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
               Height          =   255
               Left            =   150
               TabIndex        =   98
               Top             =   180
               Value           =   -1  'True
               Width           =   1845
            End
            Begin VB.OptionButton optNonGenuine 
               Caption         =   "Non-Genuine"
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
               Height          =   255
               Left            =   150
               TabIndex        =   97
               Top             =   420
               Width           =   1845
            End
         End
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   74
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":0D47
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":0E99
         Style           =   1  'Graphical
         TabIndex        =   72
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
         ItemData        =   "frmPMISTrans_ADB_Issuances.frx":11D7
         Left            =   1440
         List            =   "frmPMISTrans_ADB_Issuances.frx":11D9
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
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   480
         Width           =   585
      End
      Begin VB.Frame fraCostToCost 
         Height          =   405
         Left            =   2190
         TabIndex        =   109
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
            TabIndex        =   110
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":11DB
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":132D
         Style           =   1  'Graphical
         TabIndex        =   73
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
         Left            =   6780
         TabIndex        =   124
         Top             =   1080
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
         Left            =   6720
         TabIndex        =   122
         Top             =   660
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
         Left            =   6810
         TabIndex        =   121
         Top             =   1440
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
         Left            =   6780
         TabIndex        =   119
         Top             =   270
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
         TabIndex        =   56
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
         TabIndex        =   54
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   55
         Top             =   1860
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9615
      ScaleHeight     =   885
      ScaleWidth      =   2220
      TabIndex        =   76
      Top             =   5775
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":167D
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":17CF
         Style           =   1  'Graphical
         TabIndex        =   77
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":1B0D
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":1C5F
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   795
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
      TabIndex        =   133
      Top             =   6660
      Width           =   2565
   End
   Begin VB.PictureBox Picture6 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11460
      TabIndex        =   112
      Top             =   6810
      Width           =   11460
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
         Left            =   2760
         TabIndex        =   136
         Top             =   0
         Width           =   1425
      End
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
         TabIndex        =   116
         Top             =   0
         Width           =   3075
      End
   End
   Begin Crystal.CrystalReport rptCustomerOrder 
      Left            =   1800
      Top             =   150
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
   Begin VB.Frame fraDetails 
      Height          =   6645
      Left            =   60
      TabIndex        =   64
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
         Top             =   390
         Value           =   -1  'True
         Width           =   2385
      End
      Begin MSComctlLib.ListView lstOrd_Hd 
         Height          =   4965
         Left            =   30
         TabIndex        =   68
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":1FAF
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
         TabIndex        =   113
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
         TabIndex        =   69
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
      Height          =   2520
      Left            =   2700
      ScaleHeight     =   2490
      ScaleWidth      =   8715
      TabIndex        =   40
      Top             =   3240
      Width           =   8745
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   30
         ScaleHeight     =   255
         ScaleWidth      =   8625
         TabIndex        =   137
         Top             =   2190
         Width           =   8655
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
            TabIndex        =   142
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
            TabIndex        =   141
            Top             =   30
            Width           =   1455
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
            TabIndex        =   140
            Top             =   30
            Width           =   1455
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
            TabIndex        =   139
            Top             =   30
            Width           =   1905
         End
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
            TabIndex        =   138
            Top             =   30
            Width           =   2445
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   8100
         Top             =   120
      End
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   2145
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3784
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
      TabIndex        =   79
      Top             =   5820
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":2111
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":2263
         Style           =   1  'Graphical
         TabIndex        =   82
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":25C9
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":271B
         Style           =   1  'Graphical
         TabIndex        =   83
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":2A81
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":2BD3
         Style           =   1  'Graphical
         TabIndex        =   89
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":2F0D
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":305F
         Style           =   1  'Graphical
         TabIndex        =   90
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":3384
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":34D6
         Style           =   1  'Graphical
         TabIndex        =   84
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":3832
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":3984
         Style           =   1  'Graphical
         TabIndex        =   85
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":3C97
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":3DE9
         Style           =   1  'Graphical
         TabIndex        =   81
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":4139
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":428B
         Style           =   1  'Graphical
         TabIndex        =   80
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":45E9
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":473B
         Style           =   1  'Graphical
         TabIndex        =   86
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":4A35
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":4B87
         Style           =   1  'Graphical
         TabIndex        =   87
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":4EDF
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":5031
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox fraSignatories 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   4980
      ScaleHeight     =   2325
      ScaleWidth      =   4380
      TabIndex        =   49
      Top             =   2130
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
         MouseIcon       =   "frmPMISTrans_ADB_Issuances.frx":5390
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":54E2
         Style           =   1  'Graphical
         TabIndex        =   75
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   90
         Width           =   1395
      End
   End
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   2700
      ScaleHeight     =   3165
      ScaleWidth      =   8685
      TabIndex        =   28
      Top             =   30
      Width           =   8715
      Begin VB.CommandButton CMD_ADD_RO 
         Caption         =   "..."
         Height          =   345
         Left            =   2760
         TabIndex        =   117
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   ".."
         Height          =   345
         Left            =   2430
         TabIndex        =   114
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmdEditTranDate 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   5190
         TabIndex        =   111
         Top             =   45
         Width           =   285
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
         Height          =   555
         Left            =   4560
         TabIndex        =   100
         Top             =   2550
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
            Top             =   180
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
            Top             =   180
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
         Left            =   8280
         TabIndex        =   71
         Top             =   30
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
         Left            =   6480
         TabIndex        =   1
         Text            =   "PIWGC06H360"
         ToolTipText     =   "Type Reference PIS Number"
         Top             =   30
         Width           =   1785
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
         Height          =   735
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
         Height          =   1065
         Left            =   60
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         ToolTipText     =   "Type complete name of customer."
         Top             =   1260
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
         Left            =   3930
         MaxLength       =   10
         TabIndex        =   3
         ToolTipText     =   "Type the date of transaction in mm/dd/yyyy format (e.g 7/5/2004)"
         Top             =   45
         Width           =   1245
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
         Top             =   915
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
         Picture         =   "frmPMISTrans_ADB_Issuances.frx":5848
         ScaleHeight     =   405
         ScaleWidth      =   435
         TabIndex        =   58
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
            TabIndex        =   59
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
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   6
         ToolTipText     =   "Input customer code (e.g. S01163)"
         Top             =   420
         Width           =   945
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
         Left            =   2820
         MaxLength       =   7
         TabIndex        =   4
         ToolTipText     =   "Type the transaction terms."
         Top             =   420
         Width           =   675
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
         Top             =   45
         Width           =   1245
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
         TabIndex        =   57
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
            TabIndex        =   62
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
            TabIndex        =   61
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
            TabIndex        =   60
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
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   5
         ToolTipText     =   "Type the transactin's RO number (e.g. A007541)"
         Top             =   830
         Width           =   1575
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
         Height          =   375
         Left            =   6690
         Locked          =   -1  'True
         TabIndex        =   115
         Top             =   90
         Visible         =   0   'False
         Width           =   1485
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
         Left            =   5670
         TabIndex        =   70
         Top             =   75
         Width           =   795
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
         TabIndex        =   63
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
         TabIndex        =   30
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
         TabIndex        =   39
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
         Left            =   4590
         TabIndex        =   38
         Top             =   480
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
         TabIndex        =   37
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
         Height          =   225
         Left            =   0
         TabIndex        =   36
         Top             =   480
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
         Left            =   2160
         TabIndex        =   35
         Top             =   450
         Width           =   555
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
         Left            =   2910
         TabIndex        =   34
         Top             =   75
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
         TabIndex        =   33
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
         Height          =   195
         Left            =   60
         TabIndex        =   32
         Top             =   120
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
         TabIndex        =   31
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
         Height          =   315
         Left            =   60
         TabIndex        =   29
         Top             =   900
         Width           =   1095
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
      Left            =   2115
      ScaleHeight     =   4665
      ScaleWidth      =   8925
      TabIndex        =   127
      Top             =   930
      Visible         =   0   'False
      Width           =   8955
      Begin MSComctlLib.ListView ListView2 
         Height          =   3795
         Left            =   30
         TabIndex        =   129
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
         TabIndex        =   135
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
         TabIndex        =   134
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
         TabIndex        =   131
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
         TabIndex        =   128
         Top             =   30
         Width           =   375
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
         TabIndex        =   132
         Top             =   30
         Width           =   675
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   405
         Left            =   0
         TabIndex        =   130
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
End
Attribute VB_Name = "frmPMISTrans_ADB_Issuances"
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
Dim rsCustomer                                         As ADODB.Recordset
Dim KCNT                                               As Integer
Dim AddorEdit                                          As String
Dim ORD_TOTUPRICE                                      As Double
Dim ORD_TOTINVAMT                                      As Double
Dim ORD_TOTVAT                                         As Double
Dim ORD_TOTQTY                                         As Double
Dim PREVORDTYPE                                        As String
Dim PREVORDNO                                          As String
Dim REPOR_STATUS                                       As String
Dim LOCALACCESS                                        As String
Private WithEvents FRM_SERIES                          As frmPMISTrans_ADB_Issuances_PISFormation
Attribute FRM_SERIES.VB_VarHelpID = -1
Private WithEvents frmAE_ADB                           As frmPMISTrans_ADB_IssuancesSearch
Attribute frmAE_ADB.VB_VarHelpID = -1
Dim LOCAL_STOCKTYPE                                    As String
Dim LOCAL_COUNTERTYPE                                  As String


Sub SETSTOCKTYPE(XXX As String)
    LOCAL_STOCKTYPE = XXX
    LOCAL_COUNTERTYPE = "RIV"
    If XXX = "P" Then
        LOCALACCESS = "PARTS ISSUANCE SERVICE ISSUANCE"
    ElseIf XXX = "A" Then
        LOCALACCESS = "ACCESSORIES SERVICE ISSUANCE"
    Else
        LOCALACCESS = "MATERIALS SERVICE ISSUANCE"
    End If

End Sub

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


    rsPRS.Open "Select tranno,refpisno from PMIS_vw_PRS WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' and SALES_ORIGIN = 'S' order by Tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not rsPRS.EOF And Not rsPRS.BOF Then
        rsPRS.MoveFirst: cboRefPRSNo.Clear
        Do While Not rsPRS.EOF
            Set rsPRS_HDDup = New ADODB.Recordset
            If LOCAL_STOCKTYPE = "P" Then
                rsPRS_HDDup.Open "SELECT REFPISNO FROM PMIS_ORD_HD WHERE TRANTYPE <> 'PRS' AND [TYPE] = 'P' AND REFPRSNO = '" & Null2String(rsPRS!REFPISNO) & "'", gconDMIS
            ElseIf LOCAL_STOCKTYPE = "A" Then
                rsPRS_HDDup.Open "SELECT REFPISNO FROM PMIS_ORD_HD WHERE TRANTYPE <> 'ARS' AND [TYPE] = 'A' AND REFPRSNO = '" & Null2String(rsPRS!REFPISNO) & "'", gconDMIS
            Else
                rsPRS_HDDup.Open "SELECT REFPISNO FROM PMIS_ORD_HD WHERE TRANTYPE <> 'MRS' AND [TYPE] = 'M' AND REFPRSNO = '" & Null2String(rsPRS!REFPISNO) & "'", gconDMIS
            End If


            If Not rsPRS_HDDup.EOF And Not rsPRS_HDDup.BOF Then

            Else
                cboRefPRSNo.AddItem Null2String(rsPRS!REFPISNO)
            End If
            rsPRS.MoveNext
        Loop
    End If
End Sub

Private Sub cboRefPRSNo_LostFocus()
    If AddorEdit = "ADD" Then
        Dim rsRR_HDDup                                 As ADODB.Recordset
        Set rsRR_HDDup = New ADODB.Recordset
        rsRR_HDDup.Open "select refpisno,tranno from PMIS_Ord_Hd where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND refprsno = '" & cboRefPRSNo.Text & "'", gconDMIS
        If Not rsRR_HDDup.EOF And Not rsRR_HDDup.BOF Then
            MsgBox "Reference Number Already Received", vbInformation, "Invalid Requistion Number"
            Exit Sub
        Else
            Set rsRR_HDDup = New ADODB.Recordset
            rsRR_HDDup.Open "select tranno,DS1,custname,custcode,rono from PMIS_vw_PRS where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND refpisno = '" & cboRefPRSNo.Text & "'", gconDMIS
            If (rsRR_HDDup.EOF Or rsRR_HDDup.BOF) Then
                MsgBox "Invalid Requisition Number!", vbInformation
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
    cboTranPartNo_Click
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
        cboTranPartNo_Click
    End If
End Sub

Private Sub Check1_Click()
    If LOCAL_STOCKTYPE = "P" Then
        If Module_Access(LOGID, "APPLY PARTS COST TO COST AMOUNT", "SYSTEM") = False Then Check1.Value = 0: Exit Sub
    ElseIf LOCAL_STOCKTYPE = "M" Then
        If Module_Access(LOGID, "APPLY ACCESSORIES COST TO COST AMOUNT", "SYSTEM") = False Then Check1.Value = 0: Exit Sub
    Else
        If Module_Access(LOGID, "APPLY MATERIALS COST TO COST AMOUNT", "SYSTEM") = False Then Check1.Value = 0: Exit Sub
    End If
    If Check1.Value = 1 Then
        txtTranUPrice.Text = txtTranUCost.Text
    Else
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
    End If
End Sub

Function CheckIfROBilled(XXX As String) As String
    Dim rsRO_DET                                       As ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("SELECT INVOICE FROM CSMS_REPOR WHERE INVOICE IS NOT NULL AND REP_OR = " & N2Str2Null(XXX))
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        CheckIfROBilled = UCase(Null2String(rsRO_DET!invoice))
    End If
    Set rsRO_DET = Nothing
End Function


Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", LOCALACCESS) = False Then Exit Sub
    AddorEdit = "ADD"
    initMemvars
    PisValidation
    Command3.Enabled = True
End Sub

Private Sub cmdAddTran_Click()
    SendToBack
    fraAddTran.Visible = True
    fraAddTran.ZOrder 0
    fraAddTran.Enabled = True
    AddorEdit = "ADD"
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
    StoreMemVars
    txtPRtranno.Visible = False
    Command3.Enabled = False
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", LOCALACCESS) = False Then Exit Sub

    On Error GoTo Errorcode:

    If MsgQuestionBox("Are you sure you want to Cancel this Transaction?", "Cancel Transaction") = True Then
        Dim PCURONHAND, PCURTISSQTY, PCURISSUANCES     As Integer
        Dim RSTDAYTRANDUP, RSPARTMASDUP                As ADODB.Recordset

        Set RSTDAYTRANDUP = New ADODB.Recordset
        RSTDAYTRANDUP.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_TdayTran where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = " & N2Str2Null(RSORD_HD!TRANTYPE), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
            RSTDAYTRANDUP.MoveFirst
            Do While Not RSTDAYTRANDUP.EOF
                Set RSPARTMASDUP = New ADODB.Recordset
                RSPARTMASDUP.Open "SELECT STOCKNO,ONHAND,TISSQTY,TISSQTY,ISSUANCES,REQSERVED,S_REQSERVED FROM PMIS_STOCKMAS WHERE STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD) & " AND TYPE='" & LOCAL_STOCKTYPE & "'", gconDMIS
                If Not RSPARTMASDUP.EOF And Not RSPARTMASDUP.BOF Then
                    PCURONHAND = N2Str2IntZero(RSPARTMASDUP!ONHAND) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                    PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!TISSQTY) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                    PCURISSUANCES = N2Str2IntZero(RSPARTMASDUP!ISSUANCES) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                    If Null2String(RSORD_HD!Status) = "P" Then
                        If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                            SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                          " REQSERVED = " & N2Str2IntZero(RSPARTMASDUP!REQSERVED) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY) & _
                                          " WHERE STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT
                            'NEW LOG AUDIT-------------------------------------------------
                            Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "PART NO: " & Null2String(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)), LOCAL_COUNTERTYPE, "")
                            'NEW LOG AUDIT-------------------------------------------------
                        Else
                            SQL_STATEMENT = "update PMIS_STOCKMAS set" & _
                                          " S_REQSERVED = " & N2Str2IntZero(RSPARTMASDUP!S_REQSERVED) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY) & _
                                          " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT
                            'NEW LOG AUDIT-------------------------------------------------
                            Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "PART NO: " & Null2String(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)), LOCAL_COUNTERTYPE, "")
                            'NEW LOG AUDIT-------------------------------------------------
                        End If
                        SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                      " ONHAND = " & PCURONHAND & "," & _
                                      " TISSQTY = " & PCURTISSQTY & "," & _
                                      " ISSUANCES = " & PCURISSUANCES & "," & _
                                      " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                                      " LASTUPDATE = '" & LOGDATE & "'" & _
                                      " WHERE STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                        gconDMIS.Execute SQL_STATEMENT
                        'NEW LOG AUDIT-------------------------------------------------
                        Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "PART NO: " & Null2String(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)), "", "")
                        'NEW LOG AUDIT-------------------------------------------------
                    End If
                    SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET" & _
                                  " STATUS = 'C'," & _
                                  " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                                  " LASTUPDATE = '" & LOGDATE & "'" & _
                                  " WHERE ID = " & RSTDAYTRANDUP!ID
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "C", LOCALACCESS, SQL_STATEMENT, labid, "Parts", txtTranNo, LOCAL_COUNTERTYPE, ""
                End If

                RSTDAYTRANDUP.MoveNext
            Loop
        End If
        SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                      " STATUS = 'C'," & _
                      " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                      " LASTUPDATE = '" & LOGDATE & "'" & _
                      " WHERE ID = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "C", LOCALACCESS, SQL_STATEMENT, labid, "Parts", txtTranNo, LOCAL_COUNTERTYPE, ""
        rsRefresh
        On Error Resume Next
        RSORD_HD.Find "id =" & labid.Caption
        StoreMemVars
    End If
    Set RSTDAYTRANDUP = Nothing
    Set RSPARTMASDUP = Nothing

    Exit Sub

Errorcode:
    ShowVBError

End Sub



Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", LOCALACCESS) = False Then Exit Sub
    AddorEdit = "EDIT"
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
    If Function_Access(LOGID, "Acess_SYSTEM", LOCALACCESS) = False Then Exit Sub
    txtTranDate.Enabled = True
    txtTranDate.Locked = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    RSORD_HD.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    RSORD_HD.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    RSORD_HD.MoveNext
    If RSORD_HD.EOF Then
        RSORD_HD.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPISNum_Click()
    Set FRM_SERIES = New frmPMISTrans_ADB_Issuances_PISFormation

    With FRM_SERIES
        .SETSTOCKTYPE (LOCAL_STOCKTYPE)
        If AddorEdit = "EDIT" Then
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


Private Sub frmAE_ADB_SETCUSTOMERINFO(XCUSTOMERCODE As String, XCUSTOMERNAME As String, XRONUMBER As String, XREMARK As String)
    txtCustCode = XCUSTOMERCODE
    txtCustName = XCUSTOMERNAME
    txtRONO = XRONUMBER

    If AddorEdit = "EDIT" Then
        If UCase(Null2String(txtRONO)) <> UCase(Null2String((RSORD_HD!RONO))) Then
            MessagePop InfoWarning, "RO number altered.", "RO Number has been altered. Items Issued to " & Null2String(RSORD_HD!RONO) & " Will be deleted During Posting."
            'gconDMIS.Execute ("DELETE FROM PMIS_TDAYTRAN WHERE TRANNO=" & N2Str2Null(rsOrd_Hd!TRANNO))
        End If
    End If
    Unload frmAE_ADB
    Set frmAE_ADB = Nothing
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
Function ComputeForADBIssuances(VTXTRONO As String, str_Stockno As String) As Long
    Dim issused_qty                                    As Long
    Dim balnces                                        As Long
    Dim RSADB                                          As ADODB.Recordset
    Set RSADB = gconDMIS.Execute("SELECT STOCK_ORD  , sum(TRANQTY) as TRANQTY FROM PMIS_ALLDAYTRAN " & _
                               " WHERE  PMIS_ALLDAYTRAN.STOCK_ORD='" & str_Stockno & "'  AND PMIS_ALLDAYTRAN.TYPE='" & LOCAL_STOCKTYPE & "' AND " & _
                               " (TRANNO  IN (SELECT TRANNO FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND RONO='" & VTXTRONO & "'   AND (STATUS='P' OR STATUS='B')) " & _
                               " OR TRANNO IN (SELECT TRANNO FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND RONO='" & VTXTRONO & "'   AND (STATUS='P' OR STATUS='B'))) " & _
                               " AND TRANTYPE='ADB' GROUP BY STOCK_ORD  ORDER BY STOCK_ORD")

    While Not RSADB.EOF
        issused_qty = GetTotal_ADB_Filled(VTXTRONO, Null2String(RSADB!STOCK_ORD))
        balnces = N2Str2Zero(RSADB!TRANQTY) - issused_qty
        RSADB.MoveNext
    Wend
    ComputeForADBIssuances = balnces
End Function

Sub ShowStock_Issused(VTXTRONO As String)
    Dim RSADB                                          As ADODB.Recordset
    Dim LST                                            As ListItem
    Dim issused_qty                                    As Long
    Dim balnces                                        As Long
    ListView2.ListItems.Clear
    If otp_RIVADB.Value = True Then
        Set RSADB = gconDMIS.Execute("SELECT STOCK_ORD ,AVG(PMIS_STOCKMAS.ONHAND) ONHAND , sum(TRANQTY) as TRANQTY FROM PMIS_ALLDAYTRAN INNER JOIN PMIS_STOCKMAS " & _
                                   " ON PMIS_STOCKMAS.TYPE=PMIS_ALLDAYTRAN.TYPE AND PMIS_ALLDAYTRAN.STOCK_ORD=PMIS_STOCKMAS.STOCKNO " & _
                                   " WHERE  PMIS_STOCKMAS.TYPE='" & LOCAL_STOCKTYPE & "' AND " & _
                                   " (TRANNO  IN (SELECT TRANNO FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND RONO='" & VTXTRONO & "'   AND (STATUS='P' OR STATUS='B')) " & _
                                   " OR TRANNO IN (SELECT TRANNO FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND RONO='" & VTXTRONO & "'   AND (STATUS='P' OR STATUS='B'))) " & _
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
                                   " ON PMIS_STOCKMAS.TYPE=PMIS_ALLDAYTRAN.TYPE AND PMIS_ALLDAYTRAN.STOCK_ORD=PMIS_STOCKMAS.STOCKNO WHERE PMIS_STOCKMAS.TYPE='" & LOCAL_STOCKTYPE & "' AND " & _
                                   " (TRANNO  IN (SELECT TRANNO FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND RONO='" & VTXTRONO & "'   AND (STATUS='P' OR STATUS='B')) " & _
                                   " OR TRANNO IN (SELECT TRANNO FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND RONO='" & VTXTRONO & "'   AND (STATUS='P' OR STATUS='B'))) " & _
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






Private Sub txtPRtranno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTranNo.Text = txtPRtranno
        txtPRtranno.Visible = False
    End If
End Sub

Private Sub cmdPost_Click()

    On Error GoTo Errorcode
    If Function_Access(LOGID, "Acess_Post", LOCALACCESS) = False Then Exit Sub

    Dim rsPrtMas                                       As New ADODB.Recordset
    Dim rsTdytran                                      As New ADODB.Recordset
    Dim blnStockremove                                 As Boolean
    Dim FILD                                           As String
    blnStockremove = False


    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text
    If FILD = "" Or FILD = "No Entry" Then
        MsgBox "Posting of Transaction without issuance of Stock(s) is not allowed.", vbCritical, "Pls. Add Line Item(s)."
        Exit Sub
    End If
    rsTdytran.Open ("select stock_ord,tranqty, ID from pmis_tdaytran where tranno = '" & txtTranNo & "' and type = '" & LOCAL_STOCKTYPE & "' and trantype in('RIV') "), gconDMIS
    If Not (rsTdytran.BOF And rsTdytran.EOF) Then
        Do While Not rsTdytran.EOF
            rsPrtMas.Open "Select STOCKNO,onhand from PMIS_STOCKMAS where STOCKNO = '" & rsTdytran!STOCK_ORD & "' ", gconDMIS
            If Not (rsPrtMas.BOF And rsPrtMas.EOF) Then
                If rsPrtMas!ONHAND <= 0 Then
                    MsgBox "Partnumber# " & rsTdytran!STOCK_ORD & " will be remove from the transaction Out of Stock"
                    SQL_STATEMENT = "delete from PMIS_TdayTran where Id = '" & rsTdytran!ID & "' "
                    gconDMIS.Execute SQL_STATEMENT
                    blnStockremove = True
                ElseIf rsPrtMas!ONHAND < rsTdytran!TRANQTY Then
                    MsgBox "SOME PARTNUMBER ONHAND IS LESS THAN YOUR REQUEST QUANTITY", vbInformation
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
    Dim PCURTISSQTY                                    As Long
    Dim PCURISSUANCES                                  As Long
    Dim RSADBISSUANCE                                  As ADODB.Recordset
    Dim ADB_ISS                                        As Long
    Dim OVERISSUANCE                                   As Boolean
    Set RSADBISSUANCE = gconDMIS.Execute("SELECT TRANQTY,STOCK_ORD FROM PMIS_TDAYTRAN WHERE TYPE='" & LOCAL_STOCKTYPE & "' AND  TRANTYPE='RIV' AND tranno =" & N2Str2Null(txtTranNo))
    OVERISSUANCE = False
    Do While Not RSADBISSUANCE.EOF
        ADB_ISS = ComputeForADBIssuances(txtRONO, Null2String(RSADBISSUANCE!STOCK_ORD))
        If ADB_ISS - N2Str2IntZero(RSADBISSUANCE!TRANQTY) < 0 Then
            OVERISSUANCE = True
            Exit Do
        End If
        RSADBISSUANCE.MoveNext
    Loop
    If OVERISSUANCE = True Then
        MsgBox "There are over issuance for this RO " & vbCrLf & "Please Check Details", vbInformation
        TXT_ROVIEW.Text = txtRONO
        otp_RIVADB.Value = True
        Command6_Click
        Command5_Click
        Exit Sub
    End If


    Dim RSTDAYTRANDUP                                  As ADODB.Recordset
    Dim RSPARTMASDUP                                   As ADODB.Recordset
    Set RSTDAYTRANDUP = New ADODB.Recordset
    RSTDAYTRANDUP.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_TdayTran where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = " & N2Str2Null(RSORD_HD!TRANTYPE), gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
        RSTDAYTRANDUP.MoveFirst
        Do While Not RSTDAYTRANDUP.EOF
            Set RSPARTMASDUP = New ADODB.Recordset
            RSPARTMASDUP.Open "SELECT STOCKNO,ONHAND,TISSQTY,ISSUANCES,REQSERVED,S_REQSERVED,NON_HARI FROM PMIS_STOCKMAS WHERE TYPE='" & LOCAL_STOCKTYPE & "' AND STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD) & "", gconDMIS
            If Not RSPARTMASDUP.EOF And Not RSPARTMASDUP.BOF Then
                PCURONHAND = N2Str2IntZero(RSPARTMASDUP!ONHAND) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!TISSQTY) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                PCURISSUANCES = N2Str2IntZero(RSPARTMASDUP!ISSUANCES) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY)

                If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                    SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                  " REQSERVED = " & N2Str2IntZero(RSPARTMASDUP!REQSERVED) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY) & _
                                  " WHERE STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD) & " AND TYPE=" & N2Str2Null(LOCAL_STOCKTYPE)
                    gconDMIS.Execute SQL_STATEMENT
                    '===================================================================
                    NEW_LogAudit "E", "PARTS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "TRAN NO: " & txtTranNo, LOCAL_COUNTERTYPE, ""
                    '===================================================================
                Else
                    SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                  " S_REQSERVED = " & N2Str2IntZero(RSPARTMASDUP!S_REQSERVED) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY) & _
                                  " WHERE STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD) & " AND TYPE=" & N2Str2Null(LOCAL_STOCKTYPE)
                    gconDMIS.Execute SQL_STATEMENT
                    '===================================================================
                    NEW_LogAudit "E", "PARTS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "TRAN NO: " & txtTranNo, LOCAL_COUNTERTYPE, ""
                    '===================================================================
                End If

                SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                              " ONHAND = " & PCURONHAND & "," & _
                              " TISSQTY = " & PCURTISSQTY & "," & _
                              " ISSUANCES = " & PCURISSUANCES & "," & _
                              " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                              " LASTUPDATE = '" & LOGDATE & "'" & _
                              " WHERE STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD) & " AND TYPE=" & N2Str2Null(LOCAL_STOCKTYPE)
                gconDMIS.Execute SQL_STATEMENT
                '===================================================================
                NEW_LogAudit "E", "PARTS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "TRAN NO: " & txtTranNo, LOCAL_COUNTERTYPE, ""
                '===================================================================
                SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET" & _
                              " STATUS = 'P'," & _
                              " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                              " LASTUPDATE = '" & LOGDATE & "'" & _
                              " WHERE ID = " & RSTDAYTRANDUP!ID
                gconDMIS.Execute SQL_STATEMENT
                '===================================================================
                NEW_LogAudit "PP", LOCALACCESS, SQL_STATEMENT, labid, "Parts", txtTranNo, LOCAL_COUNTERTYPE, ""
                '===================================================================

            End If
            RSTDAYTRANDUP.MoveNext
        Loop
    End If
    SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                  " STATUS = 'P'," & _
                  " TOTALQTY = " & ORD_TOTQTY & "," & _
                  " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                  " LASTUPDATE = '" & LOGDATE & "'" & _
                  " WHERE ID = " & labid.Caption

    gconDMIS.Execute SQL_STATEMENT
    '===================================================================
    NEW_LogAudit "P", LOCALACCESS, SQL_STATEMENT, labid, "Parts", txtTranNo, LOCAL_COUNTERTYPE, ""
    '===================================================================
    rsRefresh
    RSORD_HD.Find "id =" & labid.Caption
    StoreMemVars

    Set RSTDAYTRANDUP = Nothing
    Set RSPARTMASDUP = Nothing


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
    STR_SQLX = STR_SQLX & " (PMIS_ORD_HD.STATUS='B' OR PMIS_ORD_HD.STATUS='P') AND PMIS_ORD_HD.TYPE='" & LOCAL_STOCKTYPE & "' "
    STR_SQLX = STR_SQLX & " GROUP BY STOCK_ORD"
    STR_SQLX = STR_SQLX & " Union "

    STR_SQLX = STR_SQLX & " SELECT STOCK_ORD,SUM(TRANQTY) AS  TRANQTY  FROM PMIS_DAYTRAN "
    STR_SQLX = STR_SQLX & " INNER JOIN  PMIS_ORD_HIST ON "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TYPE=PMIS_DAYTRAN.TYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TRANTYPE=PMIS_DAYTRAN.TRANTYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TRANNO = PMIS_DAYTRAN.TRANNO "
    STR_SQLX = STR_SQLX & " WHERE PMIS_ORD_HIST.TRANTYPE='ADB' AND  PMIS_ORD_HIST.RONO='" & txtRONO & "' AND "
    STR_SQLX = STR_SQLX & " (PMIS_ORD_HIST.STATUS='B' OR PMIS_ORD_HIST.STATUS='P') AND PMIS_ORD_HIST.TYPE='" & LOCAL_STOCKTYPE & "' "
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
    End If
    ' Dim LNG As Long
    '    Dim RSTDAYTRAN As ADODB.Recordset
    '    If CheckIfROBilled(txtRONO) = "" Then
    '    Set RSTDAYTRAN = gconDMIS.Execute("SELECT STOCK_ORD, MAC FROM PMIS_TDAYTRAN WHERE TRANNO=" & N2Str2Null(txtTranNo) & " AND TYPE='P' AND TRANTYPE='RIV'")
    '    While Not RSTDAYTRAN.EOF
    ''        Call gconDMIS.Execute("UPDATE CSMS_RO_DET SET DETCOST =" & N2Str2Zero(RSTDAYTRAN!Mac) & " WHERE LIVIL='2' AND DETCDE=" & N2Str2Null(RSTDAYTRAN!STOCK_ORD), LNG)
    ''        MsgBox LNG
    ''        RSTDAYTRAN.MoveNext
    '    Wend
    '
    '    End If
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
    STR_SQLX = STR_SQLX & " (PMIS_ORD_HD.STATUS='P' OR PMIS_ORD_HD.STATUS='B') AND PMIS_ORD_HD.TYPE='" & LOCAL_STOCKTYPE & "' AND PMIS_ORD_HD.STATUS2='R' "
    STR_SQLX = STR_SQLX & " GROUP BY STOCK_ORD"
    STR_SQLX = STR_SQLX & " Union "

    STR_SQLX = STR_SQLX & " SELECT STOCK_ORD,SUM(TRANQTY) AS  TRANQTY  FROM PMIS_DAYTRAN "
    STR_SQLX = STR_SQLX & " INNER JOIN  PMIS_ORD_HIST ON "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TYPE=PMIS_DAYTRAN.TYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TRANTYPE=PMIS_DAYTRAN.TRANTYPE AND "
    STR_SQLX = STR_SQLX & " PMIS_ORD_HIST.TRANNO = PMIS_DAYTRAN.TRANNO "
    STR_SQLX = STR_SQLX & " WHERE PMIS_ORD_HIST.TRANTYPE='RIV' AND  PMIS_ORD_HIST.RONO='" & xro_no & "' AND PMIS_DAYTRAN.STOCK_ORD='" & x_stockno & "' AND "
    STR_SQLX = STR_SQLX & " (PMIS_ORD_HIST.STATUS='P' OR PMIS_ORD_HIST.STATUS='B' ) AND PMIS_ORD_HIST.TYPE='" & LOCAL_STOCKTYPE & "' AND PMIS_ORD_HIST.STATUS2='R'"
    STR_SQLX = STR_SQLX & " GROUP BY STOCK_ORD"

    Dim RSTOTAL_FILLED                                 As ADODB.Recordset
    Set RSTOTAL_FILLED = gconDMIS.Execute(STR_SQLX)
    If Not RSTOTAL_FILLED.EOF Or Not RSTOTAL_FILLED.BOF Then
        GetTotal_ADB_Filled = N2Str2Zero(RSTOTAL_FILLED!TRANQTY)
    Else
        GetTotal_ADB_Filled = 0
    End If

End Function
Private Sub cmdPrevious_Click()
    RSORD_HD.MovePrevious
    If RSORD_HD.BOF Then
        RSORD_HD.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", LOCALACCESS) = False Then Exit Sub


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


    NEW_LogAudit "V", LOCALACCESS, "", labid, "Parts", "TRAN NO: " & txtTranNo, LOCAL_COUNTERTYPE, ""

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdPrintRIV_Click()
    SERVICEPISPRINTING_BLANKFORM
    If LOCAL_STOCKTYPE = "P" Then
        Call SaveSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", txtPreparedBy)
        Call SaveSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", txtIssuedBy)
        Call SaveSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", txtApprovedBy)
    ElseIf LOCAL_STOCKTYPE = "P" Then
        Call SaveSetting("DMIS", "ACCESSORIES SERVICE ISSUANCE", "PREPARED BY", txtPreparedBy)
        Call SaveSetting("DMIS", "ACCESSORIES SERVICE ISSUANCE", "ISSUED BY", txtIssuedBy)
        Call SaveSetting("DMIS", "ACCESSORIES SERVICE ISSUANCE", "APPROVED BY", txtApprovedBy)
    Else
        Call SaveSetting("DMIS", "MATERIAL SERVICE ISSUANCE", "PREPARED BY", txtPreparedBy)
        Call SaveSetting("DMIS", "MATERIAL SERVICE ISSUANCE", "ISSUED BY", txtIssuedBy)
        Call SaveSetting("DMIS", "MATERIAL SERVICE ISSUANCE", "APPROVED BY", txtApprovedBy)
    End If
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
        MsgBox "Invalid Reference Reference Issuance Number!", vbCritical, "PIS Required!"
        Exit Sub
    End If

    If Trim(txtRONO.Text) = "" Then
        MsgBox "RO Number is Required...", vbInformation, "Pls Input RO Number..."
        Exit Sub
    End If

    If RTrim(LTrim(cboRefPRSNo.Text)) = "" Then
        MsgBox "Reference Requsition Number is Required...", vbInformation, "Pls. select PRS No."
        On Error Resume Next
        cboRefPRSNo.SetFocus
        Exit Sub
    End If

    If IsNull(txtTranNo.Text) = True Then
        MsgBox "Transaction No. must not be empty", vbInformation
        On Error Resume Next
        txtTranNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set RSFINDDUP = New ADODB.Recordset
            RSFINDDUP.Open "select trantype,tranno from PMIS_vw_ISS_HISTORY where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                MsgBox "Transaction No. already exist!", vbInformation
                txtTranNo.SetFocus
                On Error Resume Next
                Exit Sub
            End If
        Else
            If LTrim(RTrim(txtTranNo)) <> LTrim(RTrim(Null2String(RSORD_HD!TRANNO))) Then
                Set RSFINDDUP = New ADODB.Recordset
                RSFINDDUP.Open "select trantype,tranno from PMIS_vw_ISS_HISTORY where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                    MsgBox "Transaction No. already exist!", vbInformation
                    On Error Resume Next
                    txtTranNo.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

    If txtTranDate.Text = "" Or IsDate(txtTranDate.Text) = False Then
        MsgBox "Invalid Transaction Date!", vbInformation
        On Error Resume Next
        txtTranDate.SetFocus
        Exit Sub
    End If

    If AddorEdit = "EDIT" Then
        If UCase(Null2String(txtRONO)) <> UCase(Null2String((RSORD_HD!RONO))) Then
            If MsgBox("RO Number has been altered. Item(s) Issued to " & Null2String(RSORD_HD!RONO) & " Will be deleted.", vbInformation + vbYesNo) = vbNo Then
                Exit Sub
            End If
            gconDMIS.Execute ("DELETE FROM PMIS_TDAYTRAN WHERE TRANNO=" & N2Str2Null(RSORD_HD!TRANNO) & " AND TYPE='" & LOCAL_STOCKTYPE & "' AND TRANTYPE='RIV'")
        End If
    End If


    VCBOSALESMAN = N2Str2Null(cboSalesMan.Text)
    VCBOSMNAME = N2Str2Null(cboSMName.Text)

    If Left(txtTranNo.Text, 1) = "P" Then
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

    If txtremarks.Text = "Pls Type Your Message Here!" Then
        VTXTRemarks = "NULL"
    Else
        VTXTRemarks = Replace(txtremarks.Text, Chr(13), "")
        VTXTRemarks = Replace(txtremarks.Text, Chr(9), "")
        VTXTRemarks = Replace(Trim(txtremarks.Text), Chr(27), "")
        VTXTRemarks = N2Str2Null(VTXTRemarks)
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into PMIS_Ord_Hd" & _
                      " (STATUS2,TYPE,trantype,tranno,trandate,custcode,custname,chargeto,REFPRSNO,rono,rep_or,salesman,smname,terms,ttlinvamt,ds1,ds_desc1,ds_amt1,netinvamt,remarks,status,usercode,lastupdate,In_Process,REFPISNO,SALES_ORIGIN,SI_TYPE,PAY_CLASS,CHAR_YEAR,CHAR_MONTH,IS_SERIES,TRACK_CODE)" & _
                      " values ('R','" & LOCAL_STOCKTYPE & "'," & VTXTTRANTYPE & ", " & VTXTTRANNO & ", " & VTXTTRANDATE & ", " & _
                      " " & VTXTCUSTCODE & ", " & VTXTCUSTNAME & ", " & VTXTCHARGETO & "," & VTXTREFPRSNO & _
                        ", " & VTXTRONO & "," & VTXTREP_OR & ", " & VCBOSALESMAN & ", " & VCBOSMNAME & _
                        ", " & VtxtTerms & ", " & VTXTTTLINVAMT & _
                        ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                        ", " & VTXTNETINVAMT & ", " & VTXTRemarks & _
                        ", " & VStatus & ", " & Vusercode & ", " & VLastUpdate & "," & VIN_PROCESS & "," & VTXTREFERENCEPIS & ", " & XSALES_ORIGIN & ", " & XSI_TYPE & ", " & XPAY_CLASS & ", " & XCHAR_YEAR & ", " & XCHAR_MONTH & ", " & XIS_SERIES & ", " & XTRACK_CODE & ")"

        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("A", LOCALACCESS, SQL_STATEMENT, FindTransactionID(txtTranNo, "tranno", "PMIS_Ord_Hd", "DETAILS", N2Str2Null(LOCAL_STOCKTYPE), "TYPE"), "Parts", txtTranNo & " - " & VTXTREFPRSNO, LOCAL_COUNTERTYPE, "")
        ShowSuccessFullyAdded
    Else
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
                      " WHERE ID = " & labid.Caption

        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", LOCALACCESS, SQL_STATEMENT, labid, "PARTS", "TRAN NO: " & txtTranNo, LOCAL_COUNTERTYPE, "")

        SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                      " SALES_ORIGIN = " & XSALES_ORIGIN & "," & _
                      " SI_TYPE = " & XSI_TYPE & "," & _
                      " PAY_CLASS = " & XPAY_CLASS & "," & _
                      " CHAR_YEAR = " & XCHAR_YEAR & "," & _
                      " CHAR_MONTH = " & XCHAR_MONTH & "," & _
                      " IS_SERIES = " & XIS_SERIES & "," & _
                      " TRACK_CODE = " & XTRACK_CODE & "" & _
                      " WHERE ID = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", LOCALACCESS, SQL_STATEMENT, labid, "Parts", "TRAN NO: " & txtTranNo & " - " & VTXTREFPRSNO, LOCAL_COUNTERTYPE, "")
        SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET" & _
                      " TRANTYPE = " & VTXTTRANTYPE & "," & _
                      " TRANDATE = " & VTXTTRANDATE & "," & _
                      " TRANNO = " & VTXTTRANNO & _
                      " WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND TRANTYPE = '" & PREVORDTYPE & "' AND TRANNO = '" & Null2String(RSORD_HD!TRANNO) & "'"
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("EE", LOCALACCESS, SQL_STATEMENT, labid, "Parts", "TRAN NO: " & txtTranNo & " - " & VTXTREFPRSNO, LOCAL_COUNTERTYPE, "")
        ShowSuccessFullyUpdated
    End If

    If AddorEdit = "ADD" Then
        If Left(txtTranNo.Text, 1) = "P" Then
            'DO NOTHING FOR ALPHA NUMERIC SERIES
        Else
            SQL_STATEMENT = "UPDATE PMIS_COUNTER SET NEXTNUMBER = '" & NEXTCUNTER & "', LASTUPDATE = '" & LOGDATE & "', USERCODE = '" & "USER" & "' WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND MODUL = " & VTXTTRANTYPE
            gconDMIS.Execute SQL_STATEMENT
        End If
        Call NEW_LogAudit("E", "PARTS COUNTER", SQL_STATEMENT, FindTransactionID(VTXTTRANTYPE, "MODUL", "PMIS_Counter", "DETAILS", N2Str2Null("P"), "TYPE"), "", "MODUL: " & Null2String(VTXTTRANTYPE), "", "")
    Else
        SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                      " TTLINVAMT = " & ORD_TOTUPRICE & "," & _
                      " NETINVAMT = " & ORD_TOTINVAMT & _
                      " WHERE ID=" & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", LOCALACCESS, SQL_STATEMENT, labid, "", "TRAN NO: " & txtTranNo, "", "")
    End If
    fraDetails.Enabled = True
    rsRefresh
    RSORD_HD.Find "TRANNO = " & VTXTTRANNO
    cmdCancel.Value = True
    cleargrid grdDetails
    FillDetails
    cmdAddTran_Click
    X_FillSearchGrid ""
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdTranCancel_Click()
    SendToBack
    StoreMemVars
End Sub

Private Sub cmdTranDelete_Click()
    On Error GoTo Errorcode:

    If labDetID.Caption = "" Then
        ShowNothingToDeleteMsg
        Exit Sub
    End If
    If MsgQuestionBox("Delete This Line Item, Are you Sure?", "Delete Entry") = True Then
        SQL_STATEMENT = "delete from PMIS_TdayTran where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", LOCALACCESS, SQL_STATEMENT, labid, "Parts", "PART NO: " & cboTranPartNo, LOCAL_COUNTERTYPE, labDetID
        ShowDeletedMsg
    End If
    Dim CNT                                            As Integer
    Dim RSTDAYTRANDUP                                  As ADODB.Recordset
    Set RSTDAYTRANDUP = New ADODB.Recordset
    RSTDAYTRANDUP.Open "select id,itemno from PMIS_TdayTran where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND trantype = " & N2Str2Null(LOCAL_COUNTERTYPE) & " and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " order by itemno asc", gconDMIS
    If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
        RSTDAYTRANDUP.MoveFirst
        CNT = 0
        Do While Not RSTDAYTRANDUP.EOF
            CNT = CNT + 1
            SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET ITEMNO = " & Format(CNT, "0000") & " WHERE ID = " & RSTDAYTRANDUP!ID
            gconDMIS.Execute SQL_STATEMENT
            RSTDAYTRANDUP.MoveNext
        Loop
    End If
    FillDetails
    SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                  " TTLINVAMT = " & ORD_TOTUPRICE & "," & _
                  " NETINVAMT = " & ORD_TOTINVAMT & _
                  " WHERE ID = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT--------------------------------------------------------------------------------
    Call NEW_LogAudit("E", LOCALACCESS, SQL_STATEMENT, labid, "PARTS", "TRAN NO: " & txtTranNo, LOCAL_COUNTERTYPE, "")
    'NEW LOG AUDIT--------------------------------------------------------------------------------
    rsRefresh
    On Error Resume Next
    RSORD_HD.Find "id = " & labid.Caption
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
        Screen.MousePointer = 0
        Exit Sub
    End If

    If cboTranPartNo.Text = "" Then
        MessagePop InfoVoid, "Invalid Input", "Warning: Part Number must have a value"
        On Error Resume Next
        cboTranPartNo.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Dim rsTDaytranClone                            As ADODB.Recordset
        Set rsTDaytranClone = New ADODB.Recordset
        rsTDaytranClone.Open "SELECT TRANTYPE,TRANNO,ITEMNO,STOCK_ORD FROM PMIS_TDAYTRAN WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND STOCK_ORD = '" & cboTranPartNo.Text & "' AND TRANTYPE = '" & txtTranType.Text & "' AND TRANNO =" & N2Str2Null(RSORD_HD!TRANNO) & " ORDER BY ITEMNO ASC", gconDMIS
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
        Dim CURSAFESTOCK                               As Long
        Dim CURTISSQTY                                 As Long
        Dim CURRESSERVICE                              As Long
        Dim CURISSUANCES                               As Long
        Dim PREVCURORDQTY                              As Long
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "SELECT STOCKNO,ONHAND,SSTOCK,RESSERVICE,TISSQTY,ISSUANCES,MAC,NON_HARI FROM PMIS_STOCKMAS WHERE STOCKNO = '" & cboTranPartNo.Text & "' AND TYPE='" & LOCAL_STOCKTYPE & "'", gconDMIS
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            CURONHAND = N2Str2IntZero(RSPARTMAS!ONHAND)
            CURSAFESTOCK = N2Str2IntZero(RSPARTMAS!SSTOCK)
            CURTISSQTY = N2Str2IntZero(RSPARTMAS!TISSQTY)
            CURRESSERVICE = N2Str2IntZero(RSPARTMAS!RESSERVICE)
            CURISSUANCES = N2Str2IntZero(RSPARTMAS!ISSUANCES)
            ORDMAC = NumericVal(RSPARTMAS!MAC)
            If ORDMAC <= 0 Then
                ORDMAC = ComputeMacasofDate(cboTranPartNo, txtTranDate, LOCAL_STOCKTYPE)
            End If

            If ORDMAC <= 0 Then
                Screen.MousePointer = 0
                MsgBox "Warning: This Stock Number has Zero Cost! Pls Check in Stock Master File or Process Update Master File to Proceed.", vbCritical, "Stock Has Zero Cost"
                Screen.MousePointer = 0
                Exit Sub
            Else
                txtTranUCost.Text = ORDMAC
            End If
            If AddorEdit <> "ADD" Then
                PREVCURORDQTY = NumericVal(labPrevOrdQty.Caption)
                CURTISSQTY = CURTISSQTY - PREVCURORDQTY
                CURISSUANCES = CURISSUANCES - PREVCURORDQTY
            End If
            If CURONHAND <= 0 Then
                Screen.MousePointer = 0
                MsgBox "Out of Stock!", vbInformation
                Exit Sub
            End If


            If NumericVal(txtTranQty.Text) > CURONHAND Then
                Screen.MousePointer = 0
                MsgBox "Qty Ordered Exceeds Current Stock!", vbInformation
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
                Call NEW_LogAudit("MP", LOCALACCESS, CRITICAL_QUESTION, labid, "", "TRAN NO: " & txtTranNo & " " & " PART NO: " & cboTranPartNo & " " & CRITICAL_QUESTION, LOCAL_COUNTERTYPE, "")
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
                Call NEW_LogAudit("MP", LOCALACCESS, CRITICAL_QUESTION, labid, "", "TRAN NO: " & txtTranNo & " " & " PART NO: " & cboTranPartNo & " " & CRITICAL_QUESTION, LOCAL_COUNTERTYPE, "")
                MsgBox "User Action has been Log to Audit Trail", vbInformation, "Audit Trail Information"
            End If
        Else
            Screen.MousePointer = 0
            MsgBox "Warning: Issuance Unit Price for this Part Number is less than its Cost!" & vbCrLf & "                System will not allow this transaction to Proceed.", vbCritical, "Unit Price is Below Cost"
            Exit Sub
        End If
    End If


    ORDSTATUS = "'N'"

    Dim HARI_NONHARI                                   As String
    Dim rsTMP                                          As New ADODB.Recordset
    Set rsTMP = gconDMIS.Execute("Select NON_HARI from PMIS_STOCKMAS where STOCKNO = '" & cboTranPartNo.Text & "' AND TYPE='" & LOCAL_STOCKTYPE & "'")
    If Not rsTMP.EOF And Not rsTMP.BOF Then
        HARI_NONHARI = N2Str2Null(rsTMP!NON_HARI)
    Else
        HARI_NONHARI = N2Str2Null("")
    End If
    ORDIN_OUT = "'O'"
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "INSERT INTO PMIS_TDAYTRAN " & _
                        "(TYPE,TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST,MAC,TRANUPRICE,LASTUPDATE,USERCODE,STATUS,IN_OUT,NON_HARI)" & _
                      " VALUES ('" & LOCAL_STOCKTYPE & "'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                      " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                      " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                      " " & ORDTRANUCOST & "," & ORDMAC & "," & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & "," & HARI_NONHARI & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "AA", LOCALACCESS, SQL_STATEMENT, FindTransactionID(N2Str2Null(txtTranNo), "TRANNO", "PMIS_ORD_HD", "DETAILS", N2Str2Null(Null2String(ORDTRANTYPE)), "TRANTYPE"), "PARTS", "PART NO: " & cboTranPartNo, LOCAL_COUNTERTYPE, labDetID
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
        NEW_LogAudit "EE", LOCALACCESS, SQL_STATEMENT, labid, "Parts", "TRAN NO: " & txtTranNo, LOCAL_COUNTERTYPE, labDetID
        ShowSuccessFullyUpdated
    End If
    cleargrid grdDetails
    FillDetails
    SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                  " TOTALQTY = " & ORD_TOTQTY & "," & _
                  " TTLINVAMT = " & ORD_TOTUPRICE & "," & _
                  " NETINVAMT = " & ORD_TOTINVAMT & _
                  " WHERE ID = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT------------------------------------------------------------
    Call NEW_LogAudit("E", LOCALACCESS, SQL_STATEMENT, labid, "P", "TRAN NO: " & txtTranNo, "", "")
    'NEW LOG AUDIT------------------------------------------------------------

    Dim rsPRS_Header                                   As ADODB.Recordset
    Dim rsPRS_Details                                  As ADODB.Recordset
    Set rsPRS_Header = New ADODB.Recordset
    Set rsPRS_Header = gconDMIS.Execute("Select * from PMIS_vw_PRS where REFPISNO = '" & cboRefPRSNo.Text & "'")
    If Not rsPRS_Header.EOF And Not rsPRS_Header.BOF Then
        Set rsPRS_Details = New ADODB.Recordset
        Set rsPRS_Details = gconDMIS.Execute("Select * from PMIS_vw_PRS_Tran Where Tranno = " & N2Str2Null(rsPRS_Header!TRANNO) & " AND STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text))
        If Not rsPRS_Details.EOF And Not rsPRS_Details.BOF Then
            SQL_STATEMENT = "Update PMIS_vw_PRS_Tran set TRemarks = 'SERVED'  Where Tranno = " & N2Str2Null(rsPRS_Header!TRANNO) & " AND STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text)
            gconDMIS.Execute SQL_STATEMENT
            '====================================================================
            NEW_LogAudit "EE", LOCALACCESS, SQL_STATEMENT, labid, "Parts", "TRAN NO: " & txtTranNo, LOCAL_COUNTERTYPE, labDetID
            '====================================================================
        Else
        End If
    Else
    End If
    rsRefresh
    On Error Resume Next
    RSORD_HD.Find "id = " & labid.Caption
    StoreMemVars
    Screen.MousePointer = 0
    If AddorEdit = "ADD" Then
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
    Set frmAE_ADB = New frmPMISTrans_ADB_IssuancesSearch
    frmAE_ADB.SETSTOCKTYPE (LOCAL_STOCKTYPE)
    frmAE_ADB.Show 1
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
    KCNT = 0: ORD_TOTUPRICE = 0: ORD_TOTINVAMT = 0: ORD_TOTVAT = 0: ORD_TOTQTY = 0
    Dim STOCKDESCription                               As String
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND tranno = " & N2Str2Null(txtTranNo.Text) & " and trantype = " & N2Str2Null(txtTranType.Text) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        cboChargeTo.Enabled = False
        Screen.MousePointer = 11
        RSTDAYTRAN.MoveFirst
        Do While Not RSTDAYTRAN.EOF
            KCNT = KCNT + 1

            STOCKDESCription = SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_SUP))

            grdDetails.AddItem RSTDAYTRAN!ID & Chr(9) & Format(Null2String(RSTDAYTRAN!itemno), "0000") & Chr(9) & _
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


Function FillSalesMan(XXX As String) As String
    Set RSSALESMAN = New ADODB.Recordset
    RSSALESMAN.Open "select empno,signname from PMIS_vw_SalesMan where empno = '" & XXX & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSSALESMAN.EOF And Not RSSALESMAN.BOF Then
        FillSalesMan = Null2String(RSSALESMAN!signname)
        cboSalesMan.Text = Null2String(RSSALESMAN!empno)
    Else
        cboSalesMan.Text = ""
    End If
End Function


Sub X_FillSearchGrid(XXX As String)
    Dim RSORD_HD                                       As ADODB.Recordset
    lstOrd_Hd.Sorted = False
    lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False

    Set RSORD_HD = New ADODB.Recordset
    XXX = Replace(LTrim(RTrim(XXX)), "'", "")

    If optTranno.Value = True Then
        Set RSORD_HD = gconDMIS.Execute("select tranno, ID from PMIS_Ord_Hd where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND trantype = '" & LOCAL_COUNTERTYPE & "' and tranno like '" & XXX & "%' AND STATUS2='R'")
    ElseIf optRONo.Value = True Then
        Set RSORD_HD = gconDMIS.Execute("select Rono, id from PMIS_Ord_Hd where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND trantype = '" & LOCAL_COUNTERTYPE & "' and rono like '" & XXX & "%' AND STATUS2='R'  order by tranno asc")
    Else
        Set RSORD_HD = gconDMIS.Execute("select custname, id  from PMIS_Ord_Hd where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND trantype = '" & LOCAL_COUNTERTYPE & "' and CUSTNAME  like '" & XXX & "%' AND STATUS2='R'  order by CUSTNAME")
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
    Dim SQLTXT                                         As String
    Dim rsTMP                                          As New ADODB.Recordset
    Dim ISSCOUNTER                                     As Integer

    On Error GoTo Errorcode

    SQLTXT = "SELECT COUNT(*) AS BILANG FROM PMIS_vw_ISS_HISTORY WHERE (TRANTYPE = 'RIV')  AND LEFT(TRANNO,1) = '" & LOCAL_STOCKTYPE & "'"
    SQLTXT = SQLTXT & "AND [TYPE] = '" & LOCAL_STOCKTYPE & "'"

    Set rsTMP = gconDMIS.Execute(SQLTXT)
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        ISSCOUNTER = NumericVal(rsTMP!BILANG)
    End If
    ISSCOUNTER = ISSCOUNTER + 1
    txtPRtranno.Text = LOCAL_STOCKTYPE & Format(ISSCOUNTER, "00000")

    Set rsTMP = Nothing
Errorcode:
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim FILD                                           As String
    Dim PCURONHAND                                     As Integer
    Dim PCURTISSQTY                                    As Integer
    Dim PCURISSUANCES                                  As Integer
    Dim RSTDAYTRANDUP                                  As ADODB.Recordset
    Dim RSPARTMASDUP                                   As ADODB.Recordset

    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text

    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If picDetails.Visible = False Then Exit Sub

            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (Parts Issuance)"
            '====================================================================
            Call frmALL_AuditInquiry.DisplayHistory(labid, LOCALACCESS)
            '====================================================================

        Case vbKeyEscape
            If pic_viewStockADB.Visible = True Then
                pic_viewStockADB.Visible = False
            End If

            If Picture1.Visible = True Then
                SendToBack
                StoreMemVars
            End If
            If pic_viewStockADB.Visible = True Then
                pic_viewStockADB.Visible = False
            End If



            txtPRtranno.Visible = False
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(RSORD_HD!Status) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change..."
                ElseIf Null2String(RSORD_HD!Status) = "B" Then
                    MsgSpeechBox "Transactions are Already Billed-Out and cannot be Change..."
                ElseIf Null2String(RSORD_HD!Status) = "P" Then
                    MsgSpeechBox "Transactions are Already Posted and cannot be Change..."
                Else
                    cmdAddTran_Click
                    Picture1.Enabled = False
                    fraDetails.Enabled = False
                    picDetails.Enabled = False
                End If
            End If
        Case vbKeyF4
            If FILD <> "" And FILD <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(RSORD_HD!Status) <> "P" And Null2String(RSORD_HD!Status) <> "C" And Null2String(RSORD_HD!Status) <> "B" Then
                        grdDetails_DblClick
                        Picture1.Enabled = False
                        fraDetails.Enabled = False
                    End If
                End If
            End If
        Case vbKeyF5
            If FILD <> "" And FILD <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(RSORD_HD!Status) <> "P" And Null2String(RSORD_HD!Status) <> "C" And Null2String(RSORD_HD!Status) <> "B" Then
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
            If Picture1.Visible = True And Null2String(RSORD_HD!Status) = "C" Then
                If MsgBox("Are you sure you want to uncancel this transaction", vbInformation + vbYesNo) = vbYes Then
                    gconDMIS.Execute ("UPDATE PMIS_ORD_HD SET STATUS=NULL WHERE ID=" & labid)
                    rsRefresh
                    RSORD_HD.Find ("id=" & labid)
                    StoreMemVars
                End If
            End If
        Case vbKeyF12
            If Picture1.Visible = True Then
                If Null2String(RSORD_HD!Status) = "P" Then
                    If Function_Access(LOGID, "Acess_UNPost", LOCALACCESS) = False Then Exit Sub
                    Set RSTDAYTRANDUP = New ADODB.Recordset
                    RSTDAYTRANDUP.Open "SELECT ID,TRANTYPE,TRANNO,STOCK_ORD,TRANQTY FROM PMIS_TDAYTRAN WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND TRANNO = " & N2Str2Null(RSORD_HD!TRANNO) & " AND TRANTYPE = " & N2Str2Null(RSORD_HD!TRANTYPE), gconDMIS, adOpenForwardOnly, adLockReadOnly
                    If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
                        RSTDAYTRANDUP.MoveFirst
                        Do While Not RSTDAYTRANDUP.EOF
                            Set RSPARTMASDUP = New ADODB.Recordset
                            RSPARTMASDUP.Open "SELECT STOCKNO,ONHAND,TISSQTY,TISSQTY,ISSUANCES,REQSERVED,S_REQSERVED FROM PMIS_STOCKMAS WHERE STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD) & " AND TYPE='" & LOCAL_STOCKTYPE & "'", gconDMIS
                            If Not RSPARTMASDUP.EOF And Not RSPARTMASDUP.BOF Then
                                If LOCAL_COUNTERTYPE <> "ADB" Then
                                    PCURONHAND = N2Str2IntZero(RSPARTMASDUP!ONHAND) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                                    PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!TISSQTY) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                                    PCURISSUANCES = N2Str2IntZero(RSPARTMASDUP!ISSUANCES) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                                    If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then

                                        SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                                      " REQSERVED = " & N2Str2IntZero(RSPARTMASDUP!REQSERVED) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY) & _
                                                      " WHERE STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD) & " AND TYPE='" & LOCAL_STOCKTYPE & "'"
                                        gconDMIS.Execute SQL_STATEMENT
                                        'NEW LOG AUDIT-----------------------------------------
                                        Call NEW_LogAudit("E", "PARTS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), LOCAL_STOCKTYPE, "TRAN NO : " & Null2String(RSORD_HD!TRANTYPE) & " - UNPOST", LOCAL_COUNTERTYPE, "")
                                        'NEW LOG AUDIT-----------------------------------------
                                    Else
                                        SQL_STATEMENT = "update PMIS_STOCKMAS set" & _
                                                      " S_REQSERVED = " & N2Str2IntZero(RSPARTMASDUP!S_REQSERVED) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY) & _
                                                      " WHERE STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD) & " AND TYPE='" & LOCAL_STOCKTYPE & "'"
                                        gconDMIS.Execute SQL_STATEMENT
                                        'NEW LOG AUDIT-----------------------------------------
                                        Call NEW_LogAudit("E", "PARTS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "P", "TRAN NO : " & Null2String(RSORD_HD!TRANTYPE & " - UNPOST"), LOCAL_COUNTERTYPE, "")
                                        'NEW LOG AUDIT-----------------------------------------
                                    End If


                                    SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                                  " ONHAND = " & PCURONHAND & "," & _
                                                  " TISSQTY = " & PCURTISSQTY & "," & _
                                                  " ISSUANCES = " & PCURISSUANCES & "," & _
                                                  " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                                                  " LASTUPDATE = '" & LOGDATE & "'" & _
                                                  " WHERE STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD) & " AND TYPE='" & LOCAL_STOCKTYPE & "'"
                                    gconDMIS.Execute SQL_STATEMENT

                                    'NEW LOG AUDIT-----------------------------------------
                                    Call NEW_LogAudit("E", "PARTS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "P", "TRAN NO : " & Null2String(RSORD_HD!TRANTYPE) & " - UNPOST", LOCAL_COUNTERTYPE, "")
                                    'NEW LOG AUDIT-----------------------------------------
                                End If

                                SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET" & _
                                              " STATUS = 'N'," & _
                                              " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                                              " LASTUPDATE = '" & LOGDATE & "'" & _
                                              " WHERE ID = " & RSTDAYTRANDUP!ID
                                gconDMIS.Execute SQL_STATEMENT
                                NEW_LogAudit "UU", LOCALACCESS, SQL_STATEMENT, labid, "Parts", "TRAN NO: " & txtTranNo, LOCAL_COUNTERTYPE, ""
                            End If
                            RSTDAYTRANDUP.MoveNext
                        Loop
                    End If
                    SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                                  " STATUS = 'N'," & _
                                  " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                                  " LASTUPDATE = '" & LOGDATE & "'" & _
                                  " WHERE ID = " & labid.Caption
                    'for adb
                    gconDMIS.Execute ("UPDATE PMIS_ORD_HD SET STATUS3=NULL WHERE RONO=" & N2Str2Null(txtRONO))
                    gconDMIS.Execute ("UPDATE PMIS_ORD_HIST SET STATUS3=NULL WHERE RONO=" & N2Str2Null(txtRONO))

                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "U", LOCALACCESS, SQL_STATEMENT, labid, "Parts", "TRAN NO: " & txtTranNo, LOCAL_COUNTERTYPE, ""
                    rsRefresh
                    RSORD_HD.Find "id =" & labid.Caption
                    StoreMemVars

                End If
                Set RSTDAYTRANDUP = Nothing
                Set RSPARTMASDUP = Nothing
            End If

        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    CenterMe frmMain, Me, 1
    PMIS_ORDER_SHOW = True
    'Picture5.ZOrder 0
    optRONo.Enabled = False
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    initMemvars
    FillCboSalesMan
    rsRefresh
    On Error Resume Next
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then RSORD_HD.MoveLast
    StoreMemVars
    Screen.MousePointer = 0
    X_FillSearchGrid ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PMIS_ORDER_SHOW = False: Set frmPMISTrans_CustomerOrder = Nothing
    LOCAL_COUNTERTYPE = ""
    UnloadForm Me
End Sub

Private Sub grdDetails_DblClick()
    Dim FILD                                           As String
    If Null2String(RSORD_HD!Status) = "C" Then
        MsgSpeech "Transactions are Already Cancelled and cannot be Change"

        MsgBox "Transactions are Already Cancelled" & vbCrLf & _
               "and cannot be Change", vbInformation, "Edit Not Allowed!"

    ElseIf Null2String(RSORD_HD!Status) = "B" Then
        MsgSpeech "Transactions are Already Billed-Out and cannot be Change"

        MsgBox "Transactions are Already Billed-Out" & vbCrLf & _
               "and Cannot be Changed", vbInformation, "Edit Not Allowed!"
    ElseIf Null2String(RSORD_HD!Status) = "P" Then
        MsgSpeech "Transactions are Already Posted and cannot be Change"
        MsgBox "Transactions are Already Posted" & vbCrLf & _
               "and Cannot Be Changed!", vbInformation, "Edit Not Allowed!"
    Else
        grdDetails.Row = grdDetails.Row
        grdDetails.Col = 0
        FILD = grdDetails.Text
        If FILD <> "" And FILD <> "No Entry" Then
            AddorEdit = "EDIT"
            cmdTranDelete.Enabled = True
            BringToFront
            StorePartsEntry (FILD)
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
                 " ON PMIS_STOCKMAS.TYPE=PMIS_ALLDAYTRAN.TYPE AND PMIS_ALLDAYTRAN.STOCK_ORD=PMIS_STOCKMAS.STOCKNO WHERE PMIS_STOCKMAS.TYPE='" & LOCAL_STOCKTYPE & "' AND " & _
                 " (TRANNO  IN (SELECT TRANNO FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND RONO='" & txtRONO & "' AND ISNULL(STATUS2,'')  <>'R' AND STATUS='P') " & _
                 " OR TRANNO IN (SELECT TRANNO FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND RONO='" & txtRONO & "' AND ISNULL(STATUS2,'')  <>'R' AND STATUS='P'))   " & _
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

Private Sub initMemvars()

    If LOCAL_COUNTERTYPE = "RIV" Then
        Set RSCUNTER = New ADODB.Recordset
        RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND modul = 'RIV'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSCUNTER.EOF And Not RSCUNTER.BOF Then
            txtTranNo.Text = Format(Null2String(RSCUNTER!nextnumber), "000000")
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
    txtremarks.Text = "Pls Type Your Message Here!"
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

    labTranUCost.Visible = False: txtTranUCost.Visible = False

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

Private Sub lstOrd_Hd_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    On Error Resume Next
    Clipboard.SetText lstOrd_Hd.SelectedItem.Text
    RSORD_HD.MoveFirst
    RSORD_HD.Find ("ID=" & ITEM.ListSubItems(1).Text)
    StoreMemVars
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
    RSORD_HD.Open "select * from PMIS_Ord_Hd where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND trantype = 'RIV' AND STATUS2='R' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
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

        If RSORD_HD!TRANTYPE = "RIV" Then
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Parts.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'RIV' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Parts.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'ADB' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        End If
    Else
        If RSORD_HD!TRANTYPE = "RIV" Then
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
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Repair Order Number:&nbsp;</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & Null2String(RSORD_HD!RONO) & "</b></i></u></FONT></td>"
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
                Print #1, "<td width=5%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!itemno) & "</FONT></td>"
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
    RSREPOR.Open "select rep_or,niym,acct_no,invoice,plate_no,dte_rel from CSMS_repor where rep_or = '" & txtRONO.Text & "'", gconDMIS


    If Not RSREPOR.EOF And Not RSREPOR.BOF Then
        If Null2String(RSREPOR!invoice) <> "" Then
            REPOR_STATUS = "Billed-Out"
        End If
        txtCustName.Text = Null2String(RSREPOR!niym)
        txtCustCode.Text = Null2String(RSREPOR!ACCT_NO)

        If Null2String(RSREPOR!PLATE_NO) <> "" Then
            Set RSCUSTINFO = New ADODB.Recordset
            'IT MIGHT GIVE A WRONG INFO OF THE CUSTOMER
            Set RSCUSTINFO = gconDMIS.Execute("select * from CSMS_CUSVEH where Plate_NO=" & N2Str2Null(RSREPOR!PLATE_NO) & "  And CUSCDE='" & Null2String(RSREPOR!ACCT_NO) & "'")
            If Not RSCUSTINFO.EOF Or Not RSCUSTINFO.BOF Then
                txtremarks = "MODEL: " & Null2String(RSCUSTINFO("model")) & vbCrLf & "ENGINE#:" & Null2String(RSCUSTINFO("SERIAL")) & vbCrLf & "VIN#:" & Null2String(RSCUSTINFO("vin")) & vbCrLf & "PLATE#:" & Null2String(RSCUSTINFO("plate_no"))
            End If
        End If
    Else
        txtCustName.Text = ""
        txtCustCode.Text = ""

    End If
End Sub

Sub SetCustomer()
    Dim rsCustomer                                     As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer where CusCde = '" & txtCustCode.Text & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        txtCustName.Text = Null2String(rsCustomer!AcctName) & vbCrLf & Null2String(rsCustomer!CUSTOMERADD) & vbCrLf & Null2String(rsCustomer!City)
    Else
        txtCustName = ""
    End If
End Sub

Sub SetPartDetails(XXX As String)
    Dim RSPARTMAS                                      As ADODB.Recordset
    Dim STR_SQLX                                       As String
    Dim RSADBISSUE                                     As ADODB.Recordset
    Dim RSADB_FILL                                     As ADODB.Recordset
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select * from PMIS_StockMas where TYPE = '" & LOCAL_STOCKTYPE & "' and StockNo = '" & XXX & "' AND ACTIVE = 'Y'")
    TXT_ADB_ISSUE = 0
    TXT_ADB_FILL = 0
    TXT_AVL_ISSUE = 0
    TXT_CURRONHAND = 0

    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then

        If N2Str2Zero(RSPARTMAS!ONHAND) > 0 Then chkAvailableOnStock.Value = 1 Else chkAvailableOnStock.Value = 0

        TXT_CURRONHAND = N2Str2Zero(RSPARTMAS!ONHAND)

        STR_SQLX = "SELECT 'CURRENT', SUM(TRANQTY) AS TQTY FROM PMIS_TDAYTRAN WHERE TRANNO  IN"
        STR_SQLX = STR_SQLX & "(SELECT TRANNO FROM PMIS_ORD_HD WHERE TYPE='" & LOCAL_STOCKTYPE & "' AND  TRANTYPE='ADB' AND RONO='" & txtRONO & "' AND ISNULL(STATUS2,'')  <>'R' AND (STATUS='B' OR STATUS='P'))"
        STR_SQLX = STR_SQLX & "AND TRANTYPE='ADB'"
        STR_SQLX = STR_SQLX & "AND STOCK_ORD='" & XXX & "'"
        STR_SQLX = STR_SQLX & "GROUP BY STOCK_ORD"
        STR_SQLX = STR_SQLX & " Union " & vbCrLf
        STR_SQLX = STR_SQLX & "SELECT 'OLD', SUM(TRANQTY) AS TQTY FROM PMIS_DAYTRAN WHERE TRANNO  IN"
        STR_SQLX = STR_SQLX & "(SELECT TRANNO FROM PMIS_ORD_HIST WHERE TYPE='" & LOCAL_STOCKTYPE & "' AND TRANTYPE='ADB' AND RONO='" & txtRONO & "' AND ISNULL(STATUS2,'')  <>'R' AND (STATUS='B' OR STATUS='P'))"
        STR_SQLX = STR_SQLX & "AND TRANTYPE='ADB'"
        STR_SQLX = STR_SQLX & "AND STOCK_ORD='" & XXX & "'"
        STR_SQLX = STR_SQLX & "GROUP BY STOCK_ORD"
        Set RSADBISSUE = gconDMIS.Execute(STR_SQLX)

        While Not RSADBISSUE.EOF
            TXT_ADB_ISSUE = N2Str2Zero(TXT_ADB_ISSUE) + N2Str2Zero(RSADBISSUE!TQTY)
            RSADBISSUE.MoveNext
        Wend

        STR_SQLX = "SELECT 'CURRENT', SUM(TRANQTY) AS TQTY FROM PMIS_TDAYTRAN WHERE TRANNO  IN"
        STR_SQLX = STR_SQLX & "(SELECT TRANNO FROM PMIS_ORD_HD WHERE TYPE='" & LOCAL_STOCKTYPE & "' AND  TRANTYPE='RIV' AND RONO='" & txtRONO & "' AND ISNULL(STATUS2,'')  ='R' AND (STATUS='B' OR STATUS='P'))"
        STR_SQLX = STR_SQLX & "AND TRANTYPE='RIV'"
        STR_SQLX = STR_SQLX & "AND STOCK_ORD='" & XXX & "'"
        STR_SQLX = STR_SQLX & "GROUP BY STOCK_ORD"
        STR_SQLX = STR_SQLX & " Union " & vbCrLf
        STR_SQLX = STR_SQLX & "SELECT 'OLD', SUM(TRANQTY) AS TQTY FROM PMIS_DAYTRAN WHERE TRANNO  IN"
        STR_SQLX = STR_SQLX & "(SELECT TRANNO FROM PMIS_ORD_HIST WHERE TYPE='" & LOCAL_STOCKTYPE & "' AND TRANTYPE='RIV' AND RONO='" & txtRONO & "' AND ISNULL(STATUS2,'')  ='R' AND (STATUS='B' OR STATUS='P'))"
        STR_SQLX = STR_SQLX & "AND TRANTYPE='RIV'"
        STR_SQLX = STR_SQLX & "AND STOCK_ORD='" & XXX & "'"
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

        optLocalPurchase.Value = False: optImported.Value = False: optConsigned.Value = False
        optGenuine.Value = False: optNonGenuine.Value = False
        If Null2String(RSPARTMAS!PartsOrigin) = "M" Then
            optImported.Value = True
        End If
        If Null2String(RSPARTMAS!PartsOrigin) = "L" Then
            optLocalPurchase.Value = True
        End If
        If Null2String(RSPARTMAS!PartsOrigin) = "C" Then
            optConsigned.Value = True
        End If
        If Null2String(RSPARTMAS!Genuine) = "Y" Then
            optGenuine.Value = True
        Else
            optNonGenuine.Value = True
        End If
        txtModelCode.Text = Null2String(RSPARTMAS!MODELCODE)

    Else
        optLocalPurchase.Value = False
        optImported.Value = False
        optConsigned.Value = False
        optGenuine.Value = False
        optNonGenuine.Value = False
        txtModelCode.Text = ""
    End If
End Sub

Function SetPartIDSTOCKNO(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKNO from PMIS_STOCKMAS where TYPE='" & LOCAL_STOCKTYPE & "' AND STOCKNO = " & N2Str2Null(DDD) & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDSTOCKNO = Null2String(RSPARTMAS!ID)
        SetPartDetails DDD
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select srp,STOCKNO,mac,dnp from PMIS_STOCKMAS where TYPE='" & LOCAL_STOCKTYPE & "' AND STOCKNO = '" & ppp & "' AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then

            If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                SetPartPrice = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC) * 1.12)
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            Else
                SetPartPrice = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            End If

        End If
        SetPartDetails ppp
    End If
End Function

Function SetStockCost(XXX As String) As Double
    Dim rsSTKCost                                      As ADODB.Recordset
    Set rsSTKCost = New ADODB.Recordset
    Set rsSTKCost = gconDMIS.Execute("Select MAC from PMIS_STOCKMAS where TYPE='" & LOCAL_STOCKTYPE & "' AND  STOCKNO = '" & XXX & "'")
    If Not rsSTKCost.EOF And Not rsSTKCost.BOF Then
        SetStockCost = N2Str2Zero(rsSTKCost!MAC)
    End If
End Function

Function SetSTOCKDESC(ppp As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select STOCKNO,STOCKDESC,srp,mac,dnp from PMIS_STOCKMAS where TYPE='" & LOCAL_STOCKTYPE & "' AND STOCKNO= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKDESC = Null2String(RSPARTMAS!STOCKDESC)
        If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Then
            txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!dnp))
            txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
        ElseIf Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
            txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
        Else
            txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
            txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
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
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            ElseIf Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            Else
                Set rsPRS_Header = New ADODB.Recordset
                Set rsPRS_Header = gconDMIS.Execute("Select * from PMIS_vw_PRS where REFPISNO = '" & cboRefPRSNo.Text & "'")
                If Not rsPRS_Header.EOF And Not rsPRS_Header.BOF Then
                    Set rsPRS_Details = New ADODB.Recordset
                    Set rsPRS_Details = gconDMIS.Execute("Select * from PMIS_vw_PRS_Tran Where Tranno = " & N2Str2Null(rsPRS_Header!TRANNO) & " AND STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text))
                    If Not rsPRS_Details.EOF And Not rsPRS_Details.BOF Then
                        txtTranQty.Text = N2Str2Zero(rsPRS_Details!TRANQTY)
                    End If
                End If
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            End If
        Else
            txtTranUPrice.Text = "0.00"
            txtTranUCost.Text = 0
        End If
    End If
End Function


Sub StoreMemVars()
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then

        labid.Caption = RSORD_HD!ID
        txtTranType.Text = Null2String(RSORD_HD!TRANTYPE)
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
        txtRONO.Text = Null2String(RSORD_HD!RONO)
        cboSMName.Text = FillSalesMan(Null2String(RSORD_HD!salesman))
        txtTerms.Text = Null2String(RSORD_HD!Terms)
        txtTTLInvAmt.Text = ToDoubleNumber(N2Str2Zero(RSORD_HD!ttlinvamt))
        txtDS1.Text = N2Str2IntZero(RSORD_HD!ds1)
        txtDS_Desc1.Text = Null2String(RSORD_HD!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(RSORD_HD!DS_AMT1))
        txtNetInvAmt.Text = ToDoubleNumber(N2Str2Zero(RSORD_HD!NETINVAMT))
        txtremarks.Text = Null2String(RSORD_HD!REMARKS)

        If Null2String(RSORD_HD!STATUS2) = "R" Then
            LAB_ADB = "ISSUANCE AGAINST ADB"
        Else
            LAB_ADB = ""
        End If

        If Null2String(RSORD_HD!Status) = "C" Then
            labPosted.Caption = "CANCELLED"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = False
        ElseIf Null2String(RSORD_HD!Status) = "P" Then
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
        txtTranItemNo.Text = Format(Null2String(RSTDAYTRAN!itemno), "0000")
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
        X_FillSearchGrid Format(textSearch.Text, "000000")
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
    If txtremarks.Text = "Pls Type Your Message Here!" Then txtremarks.Text = ""
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub



Private Sub txtTranDate_LostFocus()
    txtTranDate.Text = Format(txtTranDate.Text, "SHORT DATE")
    'updating code:     jaa - 10292008          - Transaction Month should be equal to current month
    If IsDate(txtTranDate) = True Then
        If DateDiff("m", txtTranDate, LOGDATE) <> 0 Then
            MsgBox "Warning: Transaction Month cannot be greater or less than the current month.", vbCritical
            txtTranDate.SetFocus
        End If
    End If
End Sub

Private Sub txtTranNo_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtTranNo_LostFocus()
    txtTranNo.Text = Format(txtTranNo.Text, "000000")
    Dim RSFINDDUP                                      As ADODB.Recordset
    If AddorEdit = "ADD" Then
        Set RSFINDDUP = New ADODB.Recordset
        RSFINDDUP.Open "select trantype,tranno from PMIS_Ord_Hd where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
            MsgBox "Transaction No. already exist!", vbInformation
            On Error Resume Next
            Exit Sub
        End If
    Else
        If LTrim(RTrim(txtTranNo)) <> LTrim(RTrim(Null2String(RSORD_HD!TRANNO))) Then
            Set RSFINDDUP = New ADODB.Recordset
            RSFINDDUP.Open "select trantype,tranno from PMIS_Ord_Hd where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                MsgBox "Transaction No. already exist!", vbInformation

                Exit Sub
            End If
        End If
    End If

End Sub

Private Sub txttranQty_Change()
    If txtTranQty.Text <> "" Then
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
