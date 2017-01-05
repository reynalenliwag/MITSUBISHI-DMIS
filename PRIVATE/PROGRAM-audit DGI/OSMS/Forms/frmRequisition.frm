VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmOSMSTransactionIssuance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supply Issuance"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10995
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000F&
   Icon            =   "frmRequisition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   10995
   Begin VB.PictureBox pictureAddeditdetails 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4110
      ScaleHeight     =   795
      ScaleWidth      =   1335
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   1395
      Begin VB.CommandButton cmdAddDetail 
         Caption         =   "&Add"
         Height          =   795
         Left            =   30
         Picture         =   "frmRequisition.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   645
      End
      Begin VB.CommandButton cmdEditDetail 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   675
         Picture         =   "frmRequisition.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.Frame fraIssuanceHeader 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   2700
      TabIndex        =   5
      Top             =   -30
      Width           =   8235
      Begin VB.TextBox txtTotalAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   960
         Width           =   3105
      End
      Begin Crystal.CrystalReport rptIssuance 
         Left            =   5070
         Top             =   180
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Supplies Issuance Printout"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   960
         Width           =   3825
      End
      Begin VB.ComboBox cboIssuedBy 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   4
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtnetcount_amount 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2670
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1320
         Width           =   2025
      End
      Begin VB.ComboBox cboIssuedTo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   2
         Top             =   570
         Width           =   3855
      End
      Begin VB.TextBox txtTrans_no 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   0
         Top             =   210
         Width           =   1335
      End
      Begin VB.TextBox txtTransDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3630
         MaxLength       =   10
         TabIndex        =   1
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Issuance Amount"
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
         Height          =   315
         Left            =   4770
         TabIndex        =   21
         Top             =   630
         Width           =   3375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Dept."
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
         Height          =   285
         Left            =   90
         TabIndex        =   19
         Top             =   990
         Width           =   1005
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Issued To"
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
         Height          =   285
         Left            =   90
         TabIndex        =   18
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Count Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   2880
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Issued By"
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
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   1380
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblRRDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. Date"
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
         Height          =   345
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.Frame fraSearch 
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
      Height          =   5865
      Left            =   60
      TabIndex        =   22
      Top             =   -30
      Width           =   2595
      Begin VB.OptionButton optNum 
         Caption         =   "Transaction &Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   26
         Top             =   450
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Transaction &Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         TabIndex        =   25
         Top             =   690
         Width           =   1845
      End
      Begin VB.TextBox txtSearch 
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
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1020
         Width           =   2475
      End
      Begin MSComctlLib.ListView lstIssuance 
         Height          =   4395
         Left            =   30
         TabIndex        =   24
         Top             =   1410
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   7752
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmRequisition.frx":0EDE
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TRANSACTION DATE"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Search by:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   27
         Top             =   180
         Width           =   1065
      End
   End
   Begin VB.PictureBox picIssuanceDetails 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3885
      Left            =   3660
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   3855
      ScaleWidth      =   6480
      TabIndex        =   39
      Top             =   900
      Visible         =   0   'False
      Width           =   6510
      Begin VB.Frame fraIssuanceDetails 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Supplies Issued"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3435
         Left            =   60
         TabIndex        =   42
         Top             =   360
         Width           =   6375
         Begin VB.CommandButton cmdIDCancel 
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
            Height          =   735
            Left            =   4860
            MouseIcon       =   "frmRequisition.frx":1040
            MousePointer    =   99  'Custom
            Picture         =   "frmRequisition.frx":1192
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   2580
            Width           =   675
         End
         Begin VB.CommandButton cmdIDSave 
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
            Height          =   735
            Left            =   4200
            MouseIcon       =   "frmRequisition.frx":14D0
            MousePointer    =   99  'Custom
            Picture         =   "frmRequisition.frx":1622
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   2580
            Width           =   675
         End
         Begin VB.TextBox txtID_Serial_No 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   50
            Top             =   2200
            Width           =   1665
         End
         Begin VB.TextBox txtIDitem_No 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            TabIndex        =   49
            Top             =   330
            Width           =   1125
         End
         Begin VB.TextBox txtIDQuantity 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   48
            Top             =   1434
            Width           =   1095
         End
         Begin VB.ComboBox cboUnit 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   47
            Top             =   1817
            Width           =   1665
         End
         Begin VB.ComboBox cboSupply 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   46
            Top             =   1051
            Width           =   3735
         End
         Begin VB.ComboBox cboSupplyCode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   45
            Top             =   668
            Width           =   3765
         End
         Begin VB.TextBox txtCost 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   44
            Top             =   2583
            Width           =   1665
         End
         Begin VB.TextBox txtAmount 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            TabIndex        =   43
            Top             =   2970
            Width           =   1665
         End
         Begin VB.CommandButton cmdIDdelete 
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
            Height          =   735
            Left            =   3540
            MouseIcon       =   "frmRequisition.frx":1972
            MousePointer    =   99  'Custom
            Picture         =   "frmRequisition.frx":1AC4
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   2580
            Width           =   675
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serial No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Left            =   495
            TabIndex        =   58
            Top             =   2310
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Left            =   645
            TabIndex        =   57
            Top             =   1530
            Width           =   825
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Left            =   675
            TabIndex        =   56
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Left            =   1095
            TabIndex        =   55
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Left            =   405
            TabIndex        =   54
            Top             =   1110
            Width           =   1065
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supply Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Left            =   255
            TabIndex        =   53
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cost"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Left            =   1035
            TabIndex        =   52
            Top             =   2670
            Width           =   435
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Left            =   705
            TabIndex        =   51
            Top             =   3000
            Width           =   765
         End
      End
      Begin VB.CommandButton cmdCancelDetailProduct 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   40
         Top             =   30
         Width           =   285
      End
      Begin XtremeShortcutBar.ShortcutCaption capAccessories 
         Height          =   330
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   6495
         _Version        =   655364
         _ExtentX        =   11456
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   ":: Add Details ::"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   -2147483630
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3225
      Left            =   2700
      TabIndex        =   11
      Top             =   1770
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   5689
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorSel    =   -2147483647
      BackColorBkg    =   -2147483633
      FillStyle       =   1
      SelectionMode   =   1
      Appearance      =   0
      MousePointer    =   99
      FormatString    =   "ITEM #  | DESCRIPTION                                         |  QTY  |  UNIT      |       COST     |   AMOUNT   | ID"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmRequisition.frx":1DEF
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9480
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   36
      Top             =   5010
      Width           =   2580
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
         MouseIcon       =   "frmRequisition.frx":2109
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisition.frx":225B
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   30
         Width           =   675
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
         Left            =   90
         MouseIcon       =   "frmRequisition.frx":2599
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisition.frx":26EB
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   30
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   5190
      ScaleHeight     =   900
      ScaleWidth      =   9225
      TabIndex        =   28
      Top             =   5010
      Width           =   9225
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
         Left            =   5040
         MouseIcon       =   "frmRequisition.frx":2A3B
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisition.frx":2B8D
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   60
         Width           =   675
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
         Left            =   4380
         MouseIcon       =   "frmRequisition.frx":2EF3
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisition.frx":3045
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   60
         Width           =   675
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
         Left            =   3660
         MouseIcon       =   "frmRequisition.frx":3370
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisition.frx":34C2
         Style           =   1  'Graphical
         TabIndex        =   62
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
         Left            =   3000
         MouseIcon       =   "frmRequisition.frx":3828
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisition.frx":397A
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   60
         Width           =   675
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
         Left            =   2340
         MouseIcon       =   "frmRequisition.frx":3CD6
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisition.frx":3E28
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   60
         Width           =   675
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
         Left            =   1680
         MouseIcon       =   "frmRequisition.frx":413B
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisition.frx":428D
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   60
         Width           =   675
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
         MouseIcon       =   "frmRequisition.frx":4587
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisition.frx":46D9
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   60
         Width           =   675
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
         Left            =   360
         MouseIcon       =   "frmRequisition.frx":4A31
         MousePointer    =   99  'Custom
         Picture         =   "frmRequisition.frx":4B83
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   60
         Width           =   675
      End
   End
   Begin VB.Label labIssuanceID 
      Caption         =   "Label11"
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
      Left            =   3570
      TabIndex        =   14
      Top             =   3630
      Width           =   285
   End
   Begin VB.Label labF2 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 to Add Details"
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
      Height          =   285
      Left            =   2700
      MouseIcon       =   "frmRequisition.frx":4EE2
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   5070
      Width           =   1965
   End
   Begin VB.Label labF3 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F3 to Edit Details"
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
      Height          =   285
      Left            =   2700
      MouseIcon       =   "frmRequisition.frx":51EC
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   5430
      Width           =   1965
   End
End
Attribute VB_Name = "frmOSMSTransactionIssuance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsIssuance_Header                                  As ADODB.Recordset
Dim rsIssuance_HeaderDup                               As ADODB.Recordset
Dim rsISSUANCE_DETAILS                                 As ADODB.Recordset
Dim rsIssuance_DetailsDup                              As ADODB.Recordset
Dim rsemployee                                         As ADODB.Recordset
Dim rsSupply                                           As ADODB.Recordset
Dim rsUnit                                             As ADODB.Recordset
Dim rsSupplier                                         As ADODB.Recordset
Dim AddorEdit                                          As String
Dim PrevTransNum                                       As String
Attribute PrevTransNum.VB_VarUserMemId = 1073938440
Dim PrevItemNo                                         As String
Dim TotalISSAMount                                     As Double
Attribute TotalISSAMount.VB_VarUserMemId = 1073938443

Sub FillGrid()
    Set rsISSUANCE_DETAILS = New ADODB.Recordset

    rsISSUANCE_DETAILS.Open "select * from OSMS_ISSUANCE_DETAILS where TRANS_NO = " & N2Str2Null(txtTrans_no.Text) & " order by id_item_no asc", gconDMIS
    If Not rsISSUANCE_DETAILS.EOF And Not rsISSUANCE_DETAILS.BOF Then
        rsISSUANCE_DETAILS.MoveFirst
        cleargrid grid
        grid.ColWidth(6) = 1
        TotalISSAMount = 0
        Do While Not rsISSUANCE_DETAILS.EOF
            grid.AddItem Format(Null2String(rsISSUANCE_DETAILS!id_item_no), "0000") & Chr(9) & _
                                        SETCBOSUPPLYDescription(Null2String(rsISSUANCE_DETAILS!Supply_Code)) & Chr(9) & _
                                        N2Str2Zero(rsISSUANCE_DETAILS!ID_Quantity) & Chr(9) & _
                                        UCase(Null2String(rsISSUANCE_DETAILS!ID_Unit)) & Chr(9) & _
                                        Format(NumericVal(rsISSUANCE_DETAILS!Cost), "###,###,##0.00") & Chr(9) & _
                                        Format(N2Str2Zero(rsISSUANCE_DETAILS!ID_Quantity) * N2Str2Zero(rsISSUANCE_DETAILS!Cost), "###,###,##0.00") & Chr(9) & _
                                        rsISSUANCE_DETAILS!ID
            TotalISSAMount = TotalISSAMount + (N2Str2Zero(rsISSUANCE_DETAILS!ID_Quantity) * NumericVal(rsISSUANCE_DETAILS!Cost))
            rsISSUANCE_DETAILS.MoveNext
        Loop
        txtTotalAmount.Text = Format(TotalISSAMount, "###,###,##0.00")
        If grid.Rows > 2 Then grid.RemoveItem 1
    Else
        cleargrid grid
    End If
End Sub

Private Sub cmdAdd_Click()
    AddorEdit = "ADD"
    fraIssuanceHeader.Caption = "Add A Record"
    fraIssuanceHeader.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    grid.Enabled = False
    InitMemVarsHeader
    lstIssuance.Enabled = False
    optNum.Enabled = False
    optDate.Enabled = False
    txtSearch.Enabled = False
    
End Sub

Sub InitMemVarsHeader()
    Set rsIssuance_HeaderDup = New ADODB.Recordset
    rsIssuance_HeaderDup.Open "select trans_no from OSMS_Issuance_Header order by trans_no asc", gconDMIS
    If Not rsIssuance_HeaderDup.EOF And Not rsIssuance_HeaderDup.BOF Then
        rsIssuance_HeaderDup.MoveLast
        txtTrans_no.Text = Format(NumericVal(rsIssuance_HeaderDup!Trans_No) + 1, "000000")
    Else
        txtTrans_no.Text = "000001"
    End If
    txtTransDate.Text = Date
    txtnetcount_amount.Text = ""
    txtDept.Text = ""
    InitCBOIssuedby
    InitCBOIssuedTo
    cleargrid grid
End Sub

Sub InitMemVarsDetails()
    grid.ColWidth(6) = 1
    Set rsIssuance_DetailsDup = New ADODB.Recordset
    rsIssuance_DetailsDup.Open "select trans_no,ID_item_no from OSMS_ISSUANCE_DETAILS where trans_no = " & N2Str2Null(txtTrans_no.Text) & " order by ID_item_no asc", gconDMIS
    If Not rsIssuance_DetailsDup.EOF And Not rsIssuance_DetailsDup.BOF Then
        rsIssuance_DetailsDup.MoveLast
        txtIDitem_No = Format(NumericVal(rsIssuance_DetailsDup!id_item_no) + 1, "0000")
    Else
        txtIDitem_No = "0001"
    End If
    txtIDQuantity.Text = 1
    txtID_Serial_No.Text = ""
    txtCost.Text = "0.00"
    InitCBOSUPPLY
    InitCBOSUPPLYCODE
    InitCBOUNIT
End Sub

Sub StoreMemVarsDetails()
    Dim FieldID                                        As Long
    grid.Col = 6
    If grid.Text <> "" Then
        FieldID = grid.Text
        Set rsISSUANCE_DETAILS = New ADODB.Recordset
        rsISSUANCE_DETAILS.Open "select * from OSMS_ISSUANCE_DETAILS where ID  = " & FieldID, gconDMIS
        If Not rsISSUANCE_DETAILS.EOF And Not rsISSUANCE_DETAILS.BOF Then
            labIssuanceID.Caption = Null2String(rsISSUANCE_DETAILS!ID)
            txtIDitem_No.Text = Null2String(rsISSUANCE_DETAILS!id_item_no)
            cboSupply.Text = SETCBOSUPPLY2(Null2String(rsISSUANCE_DETAILS!Supply_Code))
            txtIDQuantity = Null2String(rsISSUANCE_DETAILS!ID_Quantity)
            cboUnit.Text = SETCBOUNIT2(Null2String(rsISSUANCE_DETAILS!ID_Unit))
            txtID_Serial_No = Null2String(rsISSUANCE_DETAILS!ID_Serial_No)
            txtCost.Text = NumericVal(rsISSUANCE_DETAILS!Cost)
        End If
    End If
End Sub



Private Sub cmdCancel_Click()
    Picture1.Visible = True
    Picture2.Visible = False
    grid.Enabled = False
    lstIssuance.Enabled = True
    txtSearch.Enabled = True
    
    optNum.Enabled = True
    optDate.Enabled = True
    txtSearch.Enabled = True
    StoreMemVarsHeader
End Sub

Private Sub cmdCancelDetailProduct_Click(Index As Integer)
    ShowHidePictureBox picIssuanceDetails.hwnd, False, Me
End Sub

Private Sub cmdDelete_Click()
    If MsgBoxXP("Are you sure you want to delete this record?", "Delete Current Record", XP_YesNo, msg_Question) = True Then
        gconDMIS.Execute "Delete from OSMS_Issuance_Header where trans_No = '" & txtTrans_no.Text & "'"
        gconDMIS.Execute "Delete from OSMS_ISSUANCE_DETAILS where trans_No = '" & txtTrans_no.Text & "'"
        rsRefresh
        StoreMemVarsHeader
    End If
End Sub

Private Sub cmdEdit_Click()
    AddorEdit = "Edit"
    fraIssuanceHeader.Enabled = True
    fraIssuanceHeader.Caption = "Edit A Record"
    On Error Resume Next
    txtTrans_no.SetFocus
    Picture1.Visible = False
    Picture2.Visible = True
    grid.Enabled = True
    PrevTransNum = txtTrans_no.Text
    optNum.Enabled = False
    optDate.Enabled = False
    txtSearch.Enabled = False
    lstIssuance.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Function RecordFound(AAA As Variant) As Boolean
    Dim rsRecordFound                                  As ADODB.Recordset
    Set rsRecordFound = New ADODB.Recordset
    Set rsRecordFound = rsIssuance_Header.Clone
    rsRecordFound.Find "trans_no = '" & AAA & "'"
    If Not rsRecordFound.EOF Then
        rsIssuance_Header.Bookmark = rsRecordFound.Bookmark
        RecordFound = True
    Else
        Set rsRecordFound = New ADODB.Recordset
        Set rsRecordFound = rsIssuance_Header.Clone
        rsRecordFound.Find "trans_Date = '" & CDate(AAA) & "'"
        If Not rsRecordFound.EOF Then
            rsIssuance_Header.Bookmark = rsRecordFound.Bookmark
            RecordFound = True
        Else
            RecordFound = False
        End If
    End If
End Function

Private Sub cmdAddDetail_Click()
    fraIssuanceDetails.Caption = "Add A Record"
    AddorEdit = "ADD"
    grid.Enabled = False
    ShowHidePictureBox picIssuanceDetails.hwnd, True, Me

    InitMemVarsDetails
    cmdIDdelete.Visible = False
    On Error Resume Next
    txtIDitem_No.SetFocus
End Sub

Private Sub cmdIDCancel_Click()
    ShowHidePictureBox picIssuanceDetails.hwnd, False, Me
    lstIssuance.Enabled = True
    StoreMemVarsHeader
    
End Sub

Private Sub cmdIDdelete_Click()
    If MsgBoxXP("Are you sure you want to delete this record?", "Delete Current Record", XP_YesNo, msg_Question) = True Then
        gconDMIS.Execute "Delete from OSMS_ISSUANCE_DETAILS where id = " & labIssuanceID.Caption
        Dim kim                                        As Integer
        Dim rsIssuanceDetails                          As ADODB.Recordset
        Set rsIssuanceDetails = New ADODB.Recordset
        rsIssuanceDetails.Open "select * from OSMS_ISSUANCE_DETAILS where TRANS_NO = " & N2Str2Null(txtTrans_no.Text) & " order by ID_Item_No asc", gconDMIS
        If Not rsIssuanceDetails.EOF And Not rsIssuanceDetails.BOF Then
            rsIssuanceDetails.MoveFirst
            kim = 0
            Do While Not rsIssuanceDetails.EOF
                kim = kim + 1
                gconDMIS.Execute "Update OSMS_Issuance_Details set ID_Item_No = '" & Format(kim, "0000") & "' where id = " & rsIssuanceDetails!ID
                rsIssuanceDetails.MoveNext
            Loop
        End If
    End If
    cmdIDCancel.Value = True
End Sub

Private Sub cmdEditDetail_Click()
    grid.Enabled = True
    grid.Col = 6
    labIssuanceID.Caption = grid.Text
    AddorEdit = "EDIT"
    StoreMemVarsDetails
    cmdIDdelete.Visible = True
    fraIssuanceDetails.Caption = "Edit A Record"
    ShowHidePictureBox picIssuanceDetails.hwnd, True, Me
    PrevItemNo = txtIDitem_No.Text
    On Error Resume Next
    txtIDitem_No.SetFocus
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    PrintSQLReport rptIssuance, OSMS_REPORT_PATH & "IssuanceRep.rpt", "{Issuance_Header.trans_no} = '" & txtTrans_no.Text & "'", OSMS_DataConn, 1
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If Picture1.Visible = True Then
                cmdAddDetail.Value = True
            End If
        Case vbKeyF3
            If Picture1.Visible = True Then
                grid_DblClick
            End If

        Case vbKeyEscape
            cmdIDCancel.Value = True
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub cmdIDSave_Click()
    Dim CheckSupply                                    As ADODB.Recordset
    Dim mysql                                          As String
    On Error GoTo ErrorHandler
    Set CheckSupply = New Recordset
    CheckSupply.Open "Select Supply_Description from OSMS_Supply where Supply_Description = '" & cboSupply.Text & "'", gconDMIS
    If cboSupplyCode.Text = "" Then
        MsgBoxXP "Pls. Enter Supply Code!", "Enter Supply Code", XP_OKOnly, msg_Information
        On Error Resume Next
        cboSupplyCode.SetFocus
        Exit Sub
    Else
        If cboSupply.Text = "" Then
            MsgBoxXP "Pls. Enter Supply Description!", "Enter Supply Description", XP_OKOnly, msg_Information
            On Error Resume Next
            cboSupply.SetFocus
            Exit Sub
        End If
    End If
    If CheckSupply.EOF And CheckSupply.BOF Then
        If MsgBoxXP("Supply Not Found. Add this to the database?", "Supply", XP_YesNo, msg_Question) = True Then
            mysql = "Insert into OSMS_Supply" & _
                    "(supply_code, supply_description) values (" & N2Str2Null(cboSupplyCode.Text) & ", " & N2Str2Null(cboSupply.Text) & ")"
            gconDMIS.Execute mysql
        Else
            Exit Sub
        End If
    End If

    If AddorEdit = "ADD" Then
        mysql = "Insert into OSMS_Issuance_Details" & _
                "(trans_no, trans_date, ID_item_no, Supply_code, ID_quantity, ID_unit, ID_Serial_No,Cost) values (" & N2Str2Null(txtTrans_no.Text) & "," & N2Str2Null(txtTransDate.Text) & ", " & N2Str2Null(txtIDitem_No.Text) & "," & N2Str2Null(SETCBOSUPPLY(cboSupply.Text)) & ", " & NumericVal(txtIDQuantity) & ", " & N2Str2Null(SETCBOUNIT(cboUnit.Text)) & ", " & N2Str2Null(txtID_Serial_No.Text) & ", " & NumericVal(txtCost.Text) & ")"
        gconDMIS.Execute mysql
    Else
        gconDMIS.Execute "Update OSMS_Issuance_Details set " & _
                         "ID_item_no = " & N2Str2Null(txtIDitem_No.Text) & "," & _
                         "Supply_Code = " & N2Str2Null(SETCBOSUPPLY(cboSupply.Text)) & "," & _
                         "ID_quantity = " & N2Str2Null(txtIDQuantity.Text) & "," & _
                         "ID_unit = " & N2Str2Null(SETCBOUNIT(cboUnit.Text)) & "," & _
                         "ID_Serial_no = " & N2Str2Null(txtID_Serial_No.Text) & "," & _
                         "Cost = " & NumericVal(txtCost.Text) & _
                       " where id = " & labIssuanceID.Caption
    End If
    gconDMIS.Execute "Update OSMS_Supply set LASTISSUEDATE = " & N2Date2Null(txtTransDate.Text) & ", ONHAND = ONHAND - " & NumericVal(txtIDQuantity.Text) & " where Supply_Code =  " & N2Str2Null(SETCBOSUPPLY(cboSupply.Text))
    gconDMIS.Execute "update OSMS_Issuance_Header set " & _
                     "trans_No = " & N2Str2Null(txtTrans_no.Text) & "," & _
                     "trans_Date = " & N2Date2Null(txtTransDate.Text) & "," & _
                     "Issued_By = " & N2Str2Null(SETCBOIssuedby(cboIssuedBy.Text)) & "," & _
                     "Issued_To = " & N2Str2Null(SETCBOIssuedTo(cboIssuedTo.Text)) & "," & _
                     "Total_Amount = " & NumericVal(txtTotalAmount.Text) & "," & _
                     "NetCount_Amt = " & NumericVal(txtnetcount_amount.Text) & _
                   " where trans_No = " & N2Str2Null(PrevTransNum)
    cmdIDCancel.Value = True
    If AddorEdit = "ADD" Then
        cmdAddDetail_Click
    End If
    Exit Sub

ErrorHandler:
    'MsgBoxXP "Error" & Err.Number & vbCrLf & "Description: " & Err.Description, "Error", XP_OKOnly, msg_Critical
    ShowVBError
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsIssuance_Header.MoveNext
    If rsIssuance_Header.EOF Then
        ShowLastRecordMsg
        rsIssuance_Header.MoveLast
    End If
    StoreMemVarsHeader
End Sub

Sub StoreMemVarsHeader()
    If Not rsIssuance_Header.EOF And Not rsIssuance_Header.BOF Then
        txtTrans_no.Text = Null2String(rsIssuance_Header!Trans_No)
        txtTransDate.Text = Null2Date(rsIssuance_Header!TRANS_DATE)
        cboIssuedBy.Text = Null2String(SETCBOIssuedby2(rsIssuance_Header!Issued_by))
        cboIssuedTo.Text = Null2String(SETCBOIssuedTo2(rsIssuance_Header!ISSUED_TO))
        txtnetcount_amount = Null2String(rsIssuance_Header!NetCount_Amt)

        FillGrid
    Else
        MsgBoxXP "Record is Empty!", "No Record", XP_OKOnly, msg_Information
        cmdAdd.Value = True
    End If
End Sub

Sub InitCBOSUPPLYCODE()
    Set rsSupply = New Recordset
    rsSupply.Open "Select SUPPLY_CODE from OSMS_Supply order by SUPPLY_CODE asc", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        rsSupply.MoveFirst
        cboSupplyCode.Clear
        Do While Not rsSupply.EOF
            cboSupplyCode.AddItem Null2String(rsSupply!Supply_Code)
            rsSupply.MoveNext
        Loop
    End If
End Sub

Function SETCBOSUPPLYCODE(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select SUPPLY_CODE, Supply_Description, COST from OSMS_Supply WHERE SUPPLY_CODE = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLYCODE = Null2String(rsSupply!Supply_Description)
        txtCost.Text = Null2String(rsSupply!Cost)
    End If
End Function

Function SETCBOSUPPLYCODE2(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select SUPPLY_CODE, Supply_Description, COST from OSMS_Supply WHERE Supply_Description = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLYCODE2 = Null2String(rsSupply!Supply_Code)
        txtCost.Text = Null2String(rsSupply!Cost)
    End If
End Function

Function SETCBOSUPPLYDESC_APPEAR(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select SUPPLY_DESCRIPTION, SUPPLY_CODE, COST from OSMS_Supply WHERE SUPPLY_CODE = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLYDESC_APPEAR = Null2String(rsSupply!Supply_Description)
        txtCost.Text = Null2String(rsSupply!Cost)
    End If
End Function

Private Sub cboSupplyCode_LostFocus()
    If cboSupplyCode <> "" Then
        If SETCBOSUPPLYDESC_APPEAR(cboSupplyCode.Text) <> "" Then
            cboSupply = SETCBOSUPPLYDESC_APPEAR(cboSupplyCode.Text)
        End If
    End If
End Sub

Private Sub cboSupplyCode_Click()
    If cboSupplyCode <> "" Then
        If SETCBOSUPPLYDESC_APPEAR(cboSupplyCode.Text) <> "" Then
            cboSupply = SETCBOSUPPLYDESC_APPEAR(cboSupplyCode.Text)
        End If
    End If
End Sub

Sub InitCBOUNIT()
    Set rsUnit = New Recordset
    rsUnit.Open "Select Unit_description from OSMS_UNIT order by Unit_description asc", gconDMIS
    If Not rsUnit.EOF And Not rsUnit.BOF Then
        rsUnit.MoveFirst
        cboUnit.Clear
        Do While Not rsUnit.EOF
            cboUnit.AddItem rsUnit!Unit_description
            rsUnit.MoveNext
        Loop
    End If
End Sub

Function SETCBOUNIT(XXX As Variant) As String
    Set rsUnit = New Recordset
    rsUnit.Open "Select Unit_Description, Unit_Code from OSMS_UNIT WHERE unit_description = '" & XXX & "'", gconDMIS
    If Not rsUnit.EOF And Not rsUnit.BOF Then
        SETCBOUNIT = rsUnit!Unit_Code
    End If
End Function

Function SETCBOUNIT2(XXX As Variant) As String
    Set rsUnit = New Recordset
    rsUnit.Open "Select Unit_Description from OSMS_UNIT WHERE Unit_Code = '" & XXX & "'", gconDMIS
    If Not rsUnit.EOF And Not rsUnit.BOF Then
        SETCBOUNIT2 = rsUnit!Unit_description
    End If
End Function

Sub InitCBOSUPPLY()
    Set rsSupply = New Recordset
    rsSupply.Open "Select Supply_description from OSMS_Supply order by Supply_description asc", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        rsSupply.MoveFirst
        cboSupply.Clear
        Do While Not rsSupply.EOF
            cboSupply.AddItem rsSupply!Supply_Description
            rsSupply.MoveNext
        Loop
    End If
End Sub

Function SETCBOSUPPLY(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select Supply_Description, SUPPLY_CODE, COST from OSMS_Supply WHERE Supply_Description = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLY = rsSupply!Supply_Code
        txtCost.Text = Null2String(rsSupply!Cost)
    End If
End Function

Function SETCBOSUPPLY2(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select Supply_Description,SUPPLY_CODE, Cost from OSMS_Supply WHERE Supply_CODE = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLY2 = rsSupply!Supply_Description
        txtCost.Text = Null2String(rsSupply!Cost)
    End If
End Function

Function SETCBOSUPPLYDescription(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select Supply_Description,SUPPLY_CODE, Cost from OSMS_Supply WHERE Supply_CODE = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLYDescription = rsSupply!Supply_Description
    End If
End Function

Function SETCBOSUPPLYCODE_APPEAR(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select SUPPLY_CODE, SUPPLY_DESCRIPTION, COST from OSMS_Supply WHERE SUPPLY_DESCRIPTION = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLYCODE_APPEAR = Null2String(rsSupply!Supply_Code)
        txtCost.Text = Null2String(rsSupply!Cost)
    End If
End Function

Private Sub cboSupply_Change()
    If cboSupply <> "" Then
        If SETCBOSUPPLYCODE_APPEAR(cboSupply.Text) <> "" Then
            cboSupplyCode = SETCBOSUPPLYCODE_APPEAR(cboSupply.Text)
        End If
    End If
End Sub

Private Sub cboSupply_Click()
    cboSupplyCode = SETCBOSUPPLYDESC_APPEAR(cboSupply.Text)
End Sub

Private Sub cmdSave_Click()
    Dim mysql                                          As String
    On Error GoTo ErrorHandler
    If IsDate(txtTransDate.Text) = False Then
        MsgBoxXP "Transaction Date is Invalid.", "Invalid Date", XP_OKOnly, msg_Information
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        rsIssuance_Header.Find "trans_No = '" & txtTrans_no.Text & "'"
        If Not rsIssuance_Header.EOF Then
            MsgBoxXP "Transaction number already exists!", "Input New Transaction Number", XP_OKOnly, msg_Information
            On Error Resume Next
            txtTrans_no.SetFocus
            Exit Sub
        End If

        mysql = "insert into OSMS_Issuance_Header " & _
                "(trans_No, trans_Date, Issued_By, Issued_To, NetCount_Amt)" & _
                "values (" & N2Str2Null(txtTrans_no.Text) & "," & N2Date2Null(txtTransDate.Text) & "," & N2Str2Null(SETCBOIssuedby(cboIssuedBy.Text)) & "," & N2Str2Null(SETCBOIssuedTo(cboIssuedTo.Text)) & "," & N2Str2Null(txtnetcount_amount.Text) & ")"
        gconDMIS.Execute mysql
        fraIssuanceHeader.Caption = "Supplies Issued Report"
    Else
        gconDMIS.Execute "update OSMS_Issuance_Header set " & _
                         "trans_No = " & N2Str2Null(txtTrans_no.Text) & "," & _
                         "trans_Date = " & N2Date2Null(txtTransDate.Text) & "," & _
                         "Issued_By = " & N2Str2Null(SETCBOIssuedby(cboIssuedBy.Text)) & "," & _
                         "Issued_To = " & N2Str2Null(SETCBOIssuedTo(cboIssuedTo.Text)) & "," & _
                         "Total_Amount = " & N2Str2Null(txtTotalAmount.Text) & "," & _
                         "NetCount_Amt = " & N2Str2Null(txtnetcount_amount.Text) & _
                       " where trans_No = " & N2Str2Null(PrevTransNum)
        fraIssuanceHeader.Caption = "Supplies Issued Report"
    End If
    rsRefresh
    rsIssuance_Header.Find "trans_No = '" & txtTrans_no.Text & "'"
    cmdCancel.Value = True
    If AddorEdit = "ADD" Then cmdAddDetail_Click
    Exit Sub
ErrorHandler:
    MsgBoxXP "Error" & Err.Number & vbCrLf & "Description: " & Err.Description, "Error", XP_OKOnly, msg_Critical
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsIssuance_Header.MovePrevious
    If rsIssuance_Header.BOF Then

        ShowFirstRecordMsg
        rsIssuance_Header.MoveFirst
    End If
    StoreMemVarsHeader
End Sub

Function SETCBOIssuedby(XXX As Variant) As String
    Set rsemployee = New Recordset
    rsemployee.Open "Select lastname ,firstname,mi,EMPLOYEE_ID from OSMS_Employee  WHERE lastname + ', ' + firstname + ' ' + mi + '.' = '" & XXX & "'", gconDMIS
    If Not rsemployee.EOF And Not rsemployee.BOF Then
        SETCBOIssuedby = rsemployee!EMPLOYEE_ID
    End If
End Function

Function SETCBOIssuedby2(XXX As Variant) As String
    Set rsemployee = New Recordset
    rsemployee.Open "Select lastname + ', ' + firstname + ' ' + mi + '.' AS NAME,EMPLOYEE_ID from OSMS_Employee  WHERE EMPLOYEE_ID = '" & XXX & "'", gconDMIS
    If Not rsemployee.EOF And Not rsemployee.BOF Then
        SETCBOIssuedby2 = rsemployee!Name
    End If
End Function

Sub InitCBOIssuedby()
    Set rsemployee = New Recordset
    rsemployee.Open "Select lastname + ', ' + firstname + ' ' + mi + '.' AS NAME from OSMS_Employee order by lastname asc", gconDMIS
    If Not rsemployee.EOF And Not rsemployee.BOF Then
        rsemployee.MoveFirst
        cboIssuedBy.Clear
        cboIssuedBy.Text = Null2String(rsemployee![Name])
        Do While Not rsemployee.EOF
            cboIssuedBy.AddItem rsemployee![Name]
            rsemployee.MoveNext
        Loop
    End If
End Sub

Function SETCBOIssuedTo(XXX As Variant) As String
    Set rsemployee = New Recordset
    rsemployee.Open "Select lastname ,firstname,mi,EMPLOYEE_ID from OSMS_Employee WHERE lastname + ', ' + firstname + ' ' + mi + '.' = '" & XXX & "'", gconDMIS
    If Not rsemployee.EOF And Not rsemployee.BOF Then
        SETCBOIssuedTo = rsemployee!EMPLOYEE_ID
    End If
End Function

Function SETCBOIssuedTo2(XXX As Variant) As String
    Set rsemployee = New Recordset
    rsemployee.Open "Select lastname + ', ' + firstname + ' ' + mi + '.' AS NAME,EMPLOYEE_ID from OSMS_Employee WHERE EMPLOYEE_ID = '" & XXX & "'", gconDMIS
    If Not rsemployee.EOF And Not rsemployee.BOF Then
        SETCBOIssuedTo2 = rsemployee!Name
    End If
End Function

Sub InitCBOIssuedTo()
    Set rsemployee = New Recordset
    rsemployee.Open "Select lastname + ', ' + firstname + ' ' + mi + '.' AS NAME from OSMS_Employee order by lastname asc", gconDMIS
    If Not rsemployee.EOF And Not rsemployee.BOF Then
        rsemployee.MoveFirst
        cboIssuedTo.Clear
        Do While Not rsemployee.EOF
            cboIssuedTo.AddItem rsemployee![Name]
            rsemployee.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    InitMemVarsHeader
    InitMemVarsDetails
    rsRefresh
    txtSearch.Text = ""
    If Not rsIssuance_Header.EOF And Not rsIssuance_Header.BOF Then
        rsIssuance_Header.MoveLast
    End If
    cmdCancel.Value = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labF2.FontUnderline = False: labF2.FontBold = False
    labF3.FontUnderline = False: labF3.FontBold = False
End Sub

Private Sub grid_DblClick()
    If grid.Text = "" Then
        cmdAddDetail.Value = True
    Else
        cmdEditDetail.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsIssuance_Header = New ADODB.Recordset
    rsIssuance_Header.Open "Select * from OSMS_Issuance_Header order by trans_No asc", gconDMIS
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: grid_DblClick
End Sub


Private Sub labF2_Click()
    If Picture1.Visible = True Then
        cmdAddDetail.Value = True
    End If
End Sub

Private Sub labF2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labF2.FontUnderline = True
    labF2.FontBold = True
End Sub

Private Sub labF3_Click()
    If Picture1.Visible = True Then
        cmdEditDetail.Value = True
    End If
End Sub

Private Sub labF3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labF3.FontUnderline = True
    labF3.FontBold = True
End Sub



Private Sub txtIDQuantity_Change()
    ComputeAmount
End Sub

Private Sub txttransDate_LostFocus()
    txtTransDate.Text = Format(txtTransDate.Text, "DD-MMM-YY")
End Sub

Private Sub txttransDate_GotFocus()
    txtTransDate.Text = Format(txtTransDate.Text, "MM/DD/YYYY")
End Sub

Private Sub txtIDQuantity_GotFocus()
    If txtIDQuantity.Text = 1 Then
        txtIDQuantity.Text = ""
    End If
End Sub

Private Sub txtIDQuantity_LostFocus()
    ComputeAmount
    If txtIDQuantity.Text = "" Then
        txtIDQuantity.Text = 1
    End If
End Sub

Function SETCBOSUPPLIER_ADDRESS(XXX As Variant) As String
    Set rsSupplier = New Recordset
    rsSupplier.Open "Select SUPPLIER_ADDRESS, SUPPLIER_NAME from Supplier WHERE SUPPLIER_NAME = '" & XXX & "'", gconDMIS
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        SETCBOSUPPLIER_ADDRESS = Null2String(rsSupplier!Supplier_Address)
    End If
End Function

Function SETEMPDEPT(XXX As Variant) As String
    Set rsemployee = New Recordset
    rsemployee.Open "Select lastname ,firstname,mi,DEPARTMENT_CODE from OSMS_Employee WHERE lastname + ', ' + firstname + ' ' + mi + '.' = '" & XXX & "'", gconDMIS
    If Not rsemployee.EOF And Not rsemployee.BOF Then
        SETEMPDEPT = rsemployee!DEPARTMENT_CODE
    End If
End Function

Private Sub cboIssuedTo_Change()
    txtDept = SETEMPDEPT(cboIssuedTo.Text)
End Sub

Private Sub cboIssuedTo_Click()
    txtDept = SETEMPDEPT(cboIssuedTo.Text)
End Sub

Sub ComputeAmount()
    txtAmount.Text = Format(Val(txtIDQuantity.Text) * Val(txtCost.Text), "###,###,##0.00")
End Sub

Private Sub lstIssuance_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsIssuance_Header.Bookmark = rsFind(rsIssuance_Header.Clone, "Trans_No", lstIssuance.SelectedItem.SubItems(1)).Bookmark
    StoreMemVarsHeader
End Sub

Private Sub lstIssuance_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstIssuance
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

Private Sub lstIssuance_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtSearch_Change()
    If optNum.Value = True Then
        If Trim(txtSearch.Text) = "" Then FillGrid1 Else FillSearchGrid1 (txtSearch.Text)
    Else
        If Trim(txtSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (txtSearch.Text)
    End If
End Sub

Sub FillGrid2()
    Dim rsIssuance_Header2                             As ADODB.Recordset
    
    lstIssuance.Sorted = False: lstIssuance.ListItems.Clear
    lstIssuance.Enabled = False
    Set rsIssuance_Header2 = New ADODB.Recordset
    Set rsIssuance_Header2 = gconDMIS.Execute("select Trans_Date,Trans_No from OSMS_Issuance_Header order by Trans_Date asc")
    If Not (rsIssuance_Header2.EOF And rsIssuance_Header2.BOF) Then
        Listview_Loadval Me.lstIssuance.ListItems, rsIssuance_Header2
        lstIssuance.Refresh
        lstIssuance.Enabled = True
    End If
    
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsIssuance_Header2                             As ADODB.Recordset
    lstIssuance.Sorted = False: lstIssuance.ListItems.Clear
    lstIssuance.Enabled = False
    Set rsIssuance_Header2 = New ADODB.Recordset
    Set rsIssuance_Header2 = gconDMIS.Execute("select Trans_Date,Trans_No from OSMS_Issuance_Header where Trans_Date like'" & XXX & "%' order by Trans_Date asc")
    If Not (rsIssuance_Header2.EOF And rsIssuance_Header2.BOF) Then
        Listview_Loadval Me.lstIssuance.ListItems, rsIssuance_Header2
        lstIssuance.Refresh
         lstIssuance.Enabled = True
    End If
   
End Sub

Sub FillGrid1()
    Dim rsIssuance_Header2                             As ADODB.Recordset
    lstIssuance.Sorted = False: lstIssuance.ListItems.Clear
    lstIssuance.Enabled = False
    Set rsIssuance_Header2 = New ADODB.Recordset
    Set rsIssuance_Header2 = gconDMIS.Execute("select Trans_No,Trans_No from OSMS_Issuance_Header order by Trans_No asc")
    If Not (rsIssuance_Header2.EOF And rsIssuance_Header2.BOF) Then
        Listview_Loadval Me.lstIssuance.ListItems, rsIssuance_Header2
        lstIssuance.Refresh
        lstIssuance.Enabled = True
    End If
    
End Sub

Sub FillSearchGrid1(XXX As String)
    Dim rsIssuance_Header2                             As ADODB.Recordset
    lstIssuance.Sorted = False: lstIssuance.ListItems.Clear
    lstIssuance.Enabled = False
    Set rsIssuance_Header2 = New ADODB.Recordset
    Set rsIssuance_Header2 = gconDMIS.Execute("select Trans_No,Trans_No from OSMS_Issuance_Header where Trans_No like'" & XXX & "%' order by Trans_No asc")
    If Not (rsIssuance_Header2.EOF And rsIssuance_Header2.BOF) Then
        Listview_Loadval Me.lstIssuance.ListItems, rsIssuance_Header2
        lstIssuance.Refresh
        lstIssuance.Enabled = True
    End If
    
End Sub

Private Sub optNum_Click()
    If txtSearch = "" Then FillGrid1 Else FillSearchGrid1 (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub
Private Sub optDate_Click()
    If txtSearch = "" Then FillGrid2 Else FillSearchGrid2 (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub



