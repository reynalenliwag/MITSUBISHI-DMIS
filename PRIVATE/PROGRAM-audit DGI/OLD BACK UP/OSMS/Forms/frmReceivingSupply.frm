VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmOSMSTransactionReceivingSupply 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receiving Supply"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   FillColor       =   &H00404040&
   ForeColor       =   &H8000000F&
   Icon            =   "frmReceivingSupply.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   10770
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3510
      ScaleHeight     =   855
      ScaleWidth      =   1395
      TabIndex        =   44
      Top             =   5550
      Visible         =   0   'False
      Width           =   1395
      Begin VB.CommandButton cmdEditDetail 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   660
         Picture         =   "frmReceivingSupply.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   0
         Width           =   675
      End
      Begin VB.CommandButton cmdAddDetail 
         Caption         =   "&Add"
         Height          =   795
         Left            =   0
         Picture         =   "frmReceivingSupply.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   0
         Width           =   675
      End
   End
   Begin Crystal.CrystalReport rptMRR 
      Left            =   7590
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Supplies Receiving Report"
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
   Begin VB.Frame fmeRRHeader 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   2640
      TabIndex        =   11
      Top             =   -30
      Width           =   8085
      Begin VB.TextBox txtInv_Date 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3540
         TabIndex        =   8
         Top             =   2340
         Width           =   1215
      End
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
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   840
         Width           =   3075
      End
      Begin VB.TextBox txtInv_No 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   990
         TabIndex        =   7
         Top             =   2340
         Width           =   1215
      End
      Begin VB.TextBox txtrrDate 
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
         Left            =   3510
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   210
         Width           =   1245
      End
      Begin VB.TextBox txtrrNumber 
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
         Left            =   1110
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "Text3"
         Top             =   210
         Width           =   1275
      End
      Begin VB.TextBox txtRemarks 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1950
         Width           =   3045
      End
      Begin VB.TextBox txtSUPPLIER_ADDRESS 
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
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "frmReceivingSupply.frx":091E
         Top             =   990
         Width           =   4635
      End
      Begin VB.TextBox txtPONo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   990
         TabIndex        =   5
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtPODate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3540
         TabIndex        =   6
         Top             =   1980
         Width           =   1215
      End
      Begin VB.ComboBox cboSupplier 
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
         Left            =   1110
         TabIndex        =   2
         Top             =   570
         Width           =   3645
      End
      Begin VB.ComboBox cboReceivedBy 
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
         Left            =   1410
         TabIndex        =   4
         Top             =   1560
         Width           =   3345
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
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
         Height          =   255
         Left            =   2310
         TabIndex        =   25
         Top             =   2370
         Width           =   1545
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total MRR Amount"
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
         Left            =   4410
         TabIndex        =   24
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice #"
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
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2370
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Purpose"
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
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   1590
         Width           =   1335
      End
      Begin VB.Label lblRRDate 
         BackStyle       =   0  'Transparent
         Caption         =   "MRR Date"
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
         Height          =   315
         Left            =   2460
         TabIndex        =   17
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "P.O. No."
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
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2010
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Received By"
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1620
         Width           =   1665
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "P.O. Date"
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
         Height          =   255
         Left            =   2580
         TabIndex        =   13
         Top             =   2010
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MRR No."
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
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame Trans_No 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6345
      Left            =   30
      TabIndex        =   26
      Top             =   0
      Width           =   2565
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
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
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1110
         Width           =   2445
      End
      Begin VB.OptionButton optDate 
         Caption         =   "MRR &Date"
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
         Left            =   300
         TabIndex        =   28
         Top             =   720
         Width           =   1845
      End
      Begin VB.OptionButton optNum 
         Caption         =   "MRR &Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   27
         Top             =   480
         Value           =   -1  'True
         Width           =   2175
      End
      Begin MSComctlLib.ListView lstReceiving 
         Height          =   4785
         Left            =   30
         TabIndex        =   30
         Top             =   1500
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   8440
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
         MouseIcon       =   "frmReceivingSupply.frx":0924
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
      Begin VB.Label Label4 
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
         Left            =   60
         TabIndex        =   31
         Top             =   210
         Width           =   1065
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2715
      Left            =   2640
      TabIndex        =   10
      Top             =   2760
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   4789
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorSel    =   -2147483647
      BackColorBkg    =   -2147483633
      FillStyle       =   1
      SelectionMode   =   1
      Appearance      =   0
      MousePointer    =   99
      FormatString    =   "ITEM NO. |   DESCRIPTION                 |   QTY   |       UNIT       |      COST      |      AMOUNT     |   ID    "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmReceivingSupply.frx":0A86
   End
   Begin VB.PictureBox picRRDetails 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3555
      Left            =   4290
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   3525
      ScaleWidth      =   5520
      TabIndex        =   47
      Top             =   1260
      Visible         =   0   'False
      Width           =   5550
      Begin VB.Frame fmeRRDetails 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Receiving Report Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3045
         Left            =   90
         TabIndex        =   50
         Top             =   360
         Width           =   5385
         Begin VB.TextBox txtitemNo 
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
            Height          =   315
            Left            =   1530
            TabIndex        =   60
            Top             =   330
            Width           =   1125
         End
         Begin VB.TextBox txtRRDQuantity 
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
            Left            =   1530
            TabIndex        =   59
            Top             =   1500
            Width           =   1005
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
            Left            =   1530
            TabIndex        =   58
            Top             =   1860
            Width           =   1545
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
            Left            =   1530
            TabIndex        =   57
            Top             =   1080
            Width           =   3765
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
            Left            =   1530
            TabIndex        =   56
            Top             =   690
            Width           =   3765
         End
         Begin VB.TextBox txtCost 
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
            Left            =   1530
            TabIndex        =   55
            Top             =   2250
            Width           =   1545
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
            Left            =   1530
            TabIndex        =   54
            Top             =   2640
            Width           =   1545
         End
         Begin VB.CommandButton cmdRRDCancel 
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
            Height          =   705
            Left            =   4620
            MouseIcon       =   "frmReceivingSupply.frx":0DA0
            MousePointer    =   99  'Custom
            Picture         =   "frmReceivingSupply.frx":0EF2
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   2280
            Width           =   645
         End
         Begin VB.CommandButton cmdrrdSave 
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
            Height          =   705
            Left            =   3990
            MouseIcon       =   "frmReceivingSupply.frx":1230
            MousePointer    =   99  'Custom
            Picture         =   "frmReceivingSupply.frx":1382
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   2280
            Width           =   645
         End
         Begin VB.CommandButton cmdRRDdelete 
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
            Height          =   705
            Left            =   3360
            MouseIcon       =   "frmReceivingSupply.frx":16D2
            MousePointer    =   99  'Custom
            Picture         =   "frmReceivingSupply.frx":1824
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   2280
            Width           =   645
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
            Left            =   690
            TabIndex        =   68
            Top             =   330
            Width           =   795
         End
         Begin VB.Label Label3 
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
            Left            =   690
            TabIndex        =   67
            Top             =   1500
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
            Left            =   1110
            TabIndex        =   66
            Top             =   1890
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
            Left            =   420
            TabIndex        =   65
            Top             =   1110
            Width           =   1065
         End
         Begin VB.Label labRRID 
            Caption         =   "Item No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1620
            TabIndex        =   64
            Top             =   750
            Width           =   1005
         End
         Begin VB.Label Label12 
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
            Left            =   270
            TabIndex        =   63
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label15 
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
            Left            =   1080
            TabIndex        =   62
            Top             =   2280
            Width           =   405
         End
         Begin VB.Label Label16 
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
            Left            =   750
            TabIndex        =   61
            Top             =   2640
            Width           =   735
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
         Left            =   5160
         TabIndex        =   48
         Top             =   30
         Width           =   285
      End
      Begin XtremeShortcutBar.ShortcutCaption capAccessories 
         Height          =   330
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   5535
         _Version        =   655364
         _ExtentX        =   9763
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
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   4980
      ScaleHeight     =   900
      ScaleWidth      =   9225
      TabIndex        =   35
      Top             =   5550
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
         MouseIcon       =   "frmReceivingSupply.frx":1B4F
         MousePointer    =   99  'Custom
         Picture         =   "frmReceivingSupply.frx":1CA1
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   0
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
         MouseIcon       =   "frmReceivingSupply.frx":2007
         MousePointer    =   99  'Custom
         Picture         =   "frmReceivingSupply.frx":2159
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   0
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
         MouseIcon       =   "frmReceivingSupply.frx":2484
         MousePointer    =   99  'Custom
         Picture         =   "frmReceivingSupply.frx":25D6
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   0
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
         MouseIcon       =   "frmReceivingSupply.frx":293C
         MousePointer    =   99  'Custom
         Picture         =   "frmReceivingSupply.frx":2A8E
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   0
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
         MouseIcon       =   "frmReceivingSupply.frx":2DEA
         MousePointer    =   99  'Custom
         Picture         =   "frmReceivingSupply.frx":2F3C
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   0
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
         MouseIcon       =   "frmReceivingSupply.frx":324F
         MousePointer    =   99  'Custom
         Picture         =   "frmReceivingSupply.frx":33A1
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   0
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
         MouseIcon       =   "frmReceivingSupply.frx":369B
         MousePointer    =   99  'Custom
         Picture         =   "frmReceivingSupply.frx":37ED
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
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
         MouseIcon       =   "frmReceivingSupply.frx":3B45
         MousePointer    =   99  'Custom
         Picture         =   "frmReceivingSupply.frx":3C97
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9300
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   32
      Top             =   5520
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
         MouseIcon       =   "frmReceivingSupply.frx":3FF6
         MousePointer    =   99  'Custom
         Picture         =   "frmReceivingSupply.frx":4148
         Style           =   1  'Graphical
         TabIndex        =   34
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
         MouseIcon       =   "frmReceivingSupply.frx":4486
         MousePointer    =   99  'Custom
         Picture         =   "frmReceivingSupply.frx":45D8
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   30
         Width           =   675
      End
   End
   Begin VB.Label labdetID 
      Caption         =   "Label11"
      Height          =   345
      Left            =   1710
      TabIndex        =   21
      Top             =   4470
      Width           =   765
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
      Left            =   2670
      MouseIcon       =   "frmReceivingSupply.frx":4928
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   5580
      Width           =   2265
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
      Left            =   2670
      MouseIcon       =   "frmReceivingSupply.frx":4C32
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   5820
      Width           =   2265
   End
End
Attribute VB_Name = "frmOSMSTransactionReceivingSupply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsrrHEADER                                         As ADODB.Recordset
Dim rsRRHeaderDup                                      As ADODB.Recordset
Dim rsrrDETAILS                                        As ADODB.Recordset
Dim rsrrDetailsDup                                     As ADODB.Recordset
Dim rsUnit                                             As ADODB.Recordset
Dim rsSupply                                           As ADODB.Recordset
Dim rsSignatories                                      As ADODB.Recordset
Dim AddorEdit                                          As String
Dim PrevRRNum                                          As String
Attribute PrevRRNum.VB_VarUserMemId = 1073938439
Dim PrevItemNo                                         As String
Dim VLast_Cost                                         As Double
Attribute VLast_Cost.VB_VarUserMemId = 1073938442
Dim TotalMRRAMount                                     As Double

Sub InitCBOSUPPLIER()
    Set rsSupply = New Recordset
    rsSupply.Open "Select Supplier_Name from OSMS_Supplier order by Supplier_Name asc", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        rsSupply.MoveFirst
        cboSupplier.Clear
        Do While Not rsSupply.EOF
            cboSupplier.AddItem Null2String(rsSupply!SUPPLIER_NAME)
            rsSupply.MoveNext
        Loop
    End If
End Sub

Sub InitCBOSUPPLY()
    Set rsSupply = New Recordset
    rsSupply.Open "Select Supply_description from OSMS_SUPPLY order by Supply_description asc", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        rsSupply.MoveFirst
        cboSupply.Clear
        Do While Not rsSupply.EOF
            cboSupply.AddItem Null2String(rsSupply!Supply_Description)
            rsSupply.MoveNext
        Loop
    End If
End Sub

Sub InitCBOSUPPLYCODE()
    Set rsSupply = New Recordset
    rsSupply.Open "Select SUPPLY_CODE from OSMS_SUPPLY order by SUPPLY_CODE asc", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        rsSupply.MoveFirst
        cboSupplyCode.Clear
        Do While Not rsSupply.EOF
            cboSupplyCode.AddItem Null2String(rsSupply!Supply_Code)
            rsSupply.MoveNext
        Loop
    End If
End Sub

Sub InitCBOUNIT()
    Set rsUnit = New Recordset
    rsUnit.Open "Select Unit_description from OSMS_UNIT order by Unit_description asc", gconDMIS
    If Not rsUnit.EOF And Not rsUnit.BOF Then
        rsUnit.MoveFirst
        cboUnit.Clear
        cboUnit.Text = Null2String(rsUnit!Unit_description)
        Do While Not rsUnit.EOF
            cboUnit.AddItem Null2String(rsUnit!Unit_description)
            rsUnit.MoveNext
        Loop
    End If
End Sub

Sub InitCBOReceivedby()
    Set rsSignatories = New Recordset
    rsSignatories.Open "Select lastname + ', ' + firstname + ' ' + mi + '.' AS NAME from OSMS_Signatories  order by lastname asc", gconDMIS
    If Not rsSignatories.EOF And Not rsSignatories.BOF Then
        rsSignatories.MoveFirst
        cboReceivedBy.Clear
        cboReceivedBy.Text = Null2String(rsSignatories![Name])
        Do While Not rsSignatories.EOF
            cboReceivedBy.AddItem Null2String(rsSignatories![Name])
            rsSignatories.MoveNext
        Loop
    End If
End Sub

Function SETCBOSUPPLY(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select Supply_Description, SUPPLY_CODE, Cost from OSMS_SUPPLY where Supply_Description = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLY = Null2String(rsSupply!Supply_Code)
        txtCost.Text = NumericVal(rsSupply!Cost)
    End If
End Function

Function SETCBOSUPPLY2(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select Supply_Description,SUPPLY_CODE, Cost from OSMS_SUPPLY where Supply_CODE = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLY2 = Null2String(rsSupply!Supply_Description)
        txtCost.Text = NumericVal(rsSupply!Cost)
    End If
End Function

Function SETCBOSUPPLYCODE_APPEAR(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select SUPPLY_CODE, SUPPLY_DESCRIPTION, Cost from OSMS_SUPPLY where SUPPLY_DESCRIPTION = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLYCODE_APPEAR = Null2String(rsSupply!Supply_Code)
        txtCost.Text = NumericVal(rsSupply!Cost)
    End If
End Function

Private Sub cboSupply_LostFocus()
    If cboSupply <> "" Then
        If SETCBOSUPPLYCODE_APPEAR(cboSupply.Text) <> "" Then
            cboSupplyCode = SETCBOSUPPLYCODE_APPEAR(cboSupply.Text)
        End If
    End If
End Sub

Private Sub cboSupply_Click()
    If cboSupply <> "" Then
        If SETCBOSUPPLYCODE_APPEAR(cboSupply.Text) <> "" Then
            cboSupplyCode = SETCBOSUPPLYCODE_APPEAR(cboSupply.Text)
        End If
    End If
End Sub

Function SETCBOSUPPLYCODE(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select SUPPLY_CODE, Supply_Description, Cost from OSMS_SUPPLY where SUPPLY_CODE = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLYCODE = Null2String(rsSupply!Supply_Description)
        txtCost.Text = NumericVal(rsSupply!Cost)
    End If
End Function

Function SETCBOSUPPLYCODE2(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select SUPPLY_CODE, Supply_Description, Cost from OSMS_SUPPLY where Supply_Description = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLYCODE2 = Null2String(rsSupply!Supply_Code)
        txtCost.Text = NumericVal(rsSupply!Cost)
    End If
End Function

Function SETCBOSUPPLYDESC_APPEAR(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select SUPPLY_DESCRIPTION, SUPPLY_CODE,COST from OSMS_SUPPLY where SUPPLY_CODE = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLYDESC_APPEAR = Null2String(rsSupply!Supply_Description)
        txtCost.Text = NumericVal(rsSupply!Cost)
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

Function SETCBOReceivedby(XXX As Variant) As String
    Set rsSignatories = New Recordset
    rsSignatories.Open "Select lastname ,firstname,mi,Signatory_ID from OSMS_Signatories  WHERE lastname + ', ' + firstname + ' ' + mi + '.' = '" & XXX & "'", gconDMIS
    If Not rsSignatories.EOF And Not rsSignatories.BOF Then
        SETCBOReceivedby = rsSignatories!SIGNATORY_ID
    End If
End Function

Function SETCBOReceivedby2(XXX As Variant) As String
    Set rsSignatories = New Recordset
    rsSignatories.Open "Select lastname + ', ' + firstname + ' ' + mi + '.' AS NAME,Signatory_ID from OSMS_Signatories  WHERE Signatory_ID = '" & XXX & "'", gconDMIS
    If Not rsSignatories.EOF And Not rsSignatories.BOF Then
        SETCBOReceivedby2 = rsSignatories!Name
    End If
End Function

Function SETCBOSUPPLIER(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select Supplier_Name,SUPPLIER_CODE from  OSMS_Supplier WHERE Supplier_Name = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLIER = Null2String(rsSupply!Supplier_code)
    End If
End Function

Function SETCBOSUPPLIER2(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select Supplier_Name,SUPPLIER_CODE from OSMS_Supplier WHERE Supplier_CODE = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLIER2 = Null2String(rsSupply!SUPPLIER_NAME)
    End If
End Function

Function SETCBOSUPPLIER_ADDRESS(XXX As Variant) As String
    Set rsSupply = New Recordset
    rsSupply.Open "Select SUPPLIER_ADDRESS, SUPPLIER_NAME from OSMS_Supplier WHERE SUPPLIER_NAME = '" & XXX & "'", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        SETCBOSUPPLIER_ADDRESS = Null2String(rsSupply!Supplier_Address)
    End If
End Function

Private Sub cboSupplier_Change()
    txtSUPPLIER_ADDRESS = SETCBOSUPPLIER_ADDRESS(cboSupplier.Text)
End Sub

Private Sub cboSupplier_Click()
    txtSUPPLIER_ADDRESS = SETCBOSUPPLIER_ADDRESS(cboSupplier.Text)
End Sub

Function SETCBOUNIT(XXX As Variant) As String
    Set rsUnit = New Recordset
    rsUnit.Open "Select Unit_Description, Unit_Code from OSMS_UNIT WHERE unit_description = '" & XXX & "'", gconDMIS
    If Not rsUnit.EOF And Not rsUnit.BOF Then
        SETCBOUNIT = Null2String(rsUnit!Unit_Code)
    End If
End Function

Function SETCBOUNIT2(XXX As Variant) As String
    Set rsUnit = New Recordset
    rsUnit.Open "Select Unit_Description from OSMS_UNIT WHERE Unit_Code = '" & XXX & "'", gconDMIS
    If Not rsUnit.EOF And Not rsUnit.BOF Then
        SETCBOUNIT2 = Null2String(rsUnit!Unit_description)
    End If
End Function

Private Sub cmdAdd_Click()
    fmeRRHeader.Caption = "Add A Record"
    AddorEdit = "ADD"
    fmeRRHeader.Enabled = True
    grid.Enabled = False
    initMemvars
    lstReceiving.Enabled = False
    txtSearch.Enabled = False
    On Error Resume Next
    cboSupplier.SetFocus
    picSaves.Visible = True
    picAdds.Visible = False
    Trans_No.Enabled = False
    Trans_No.Enabled = False
End Sub

Sub initMemvars()
    Set rsRRHeaderDup = New ADODB.Recordset
    rsRRHeaderDup.Open "select RRNUMBER from OSMS_RRHEADER  order by rrnumber asc", gconDMIS
    If Not rsRRHeaderDup.EOF And Not rsRRHeaderDup.BOF Then
        rsRRHeaderDup.MoveLast
        txtrrNumber.Text = Format(NumericVal(rsRRHeaderDup!rrnumber) + 1, "000000")
    Else
        txtrrNumber.Text = "000001"
    End If
    txtrrDate.Text = Date
    txtInv_No.Text = ""
    txtInv_Date.Text = ""
    txtPONo.Text = ""
    txtPODate.Text = ""
    txtSUPPLIER_ADDRESS.Text = ""
    txtRemarks.Text = ""
    InitCBOSUPPLIER
    InitCBOReceivedby
    cleargrid grid
End Sub

Sub InitMemVarsRRD()
    Set rsrrDetailsDup = New ADODB.Recordset
    rsrrDetailsDup.Open "select rrnumber,item_no from OSMS_RRDETAILS  where rrNumber = " & N2Str2Null(txtrrNumber.Text) & " order by item_no asc", gconDMIS
    If Not rsrrDetailsDup.EOF And Not rsrrDetailsDup.BOF Then
        rsrrDetailsDup.MoveLast
        txtitemNo = Format(NumericVal(rsrrDetailsDup!item_no) + 1, "0000")
    Else
        txtitemNo = "0001"
    End If
    txtRRDQuantity = 1
    txtCost.Text = "0.00"
    InitCBOUNIT
    InitCBOSUPPLY
    InitCBOSUPPLYCODE
End Sub

Private Sub cmdCancel_Click()
    fmeRRHeader.Enabled = False
    fmeRRHeader.Caption = ""
    grid.Enabled = False
    picSaves.Visible = False
    picAdds.Visible = True
    lstReceiving.Enabled = True
    txtSearch.Enabled = True
    Trans_No.Enabled = True
    Trans_No.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdCancelDetailProduct_Click(Index As Integer)
    ShowHidePictureBox picRRDetails.hwnd, False, Me
End Sub

Private Sub cmdDelete_Click()
    If MsgBoxXP("Are you sure you want to delete this record?", "Delete Current Record", XP_YesNo, msg_Question) = True Then
        gconDMIS.Execute "delete from OSMS_RRHEADER  where rrNumber = '" & txtrrNumber.Text & "'"
        gconDMIS.Execute "delete from OSMS_RRDETAILS  where rrNUmber = '" & txtrrNumber.Text & "'"
        rsRefresh
        StoreMemVars

        lstReceiving.ListItems.Remove (lstReceiving.SelectedItem.Index)
    End If
End Sub

Private Sub cmdEdit_Click()
    AddorEdit = "EDIT"
    fmeRRHeader.Enabled = True
    fmeRRHeader.Caption = "Edit A Record"
    grid.Enabled = True
    picSaves.Visible = True
    picAdds.Visible = False
    On Error Resume Next
    txtrrNumber.SetFocus
    PrevRRNum = txtrrNumber.Text
    Trans_No.Enabled = False
    Trans_No.Enabled = False
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
    Set rsRecordFound = rsrrHEADER.Clone
    rsRecordFound.Find "RRNumber = '" & AAA & "'"
    If Not rsRecordFound.EOF Then
        rsrrHEADER.Bookmark = rsRecordFound.Bookmark
        RecordFound = True
    Else
        Set rsRecordFound = New ADODB.Recordset
        Set rsRecordFound = rsrrHEADER.Clone
        rsRecordFound.Find "RRDate = '" & CDate(AAA) & "'"
        If Not rsRecordFound.EOF Then
            rsrrHEADER.Bookmark = rsRecordFound.Bookmark
            RecordFound = True
        Else
            RecordFound = False
        End If
    End If
End Function

Private Sub cmdNext_Click()
    On Error Resume Next
    rsrrHEADER.MoveNext
    If rsrrHEADER.EOF Then
        ShowLastRecordMsg
        rsrrHEADER.MoveLast
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsrrHEADER.MovePrevious
    If rsrrHEADER.BOF Then
        ShowFirstRecordMsg
        rsrrHEADER.MoveFirst
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    PrintSQLReport rptMRR, OSMS_REPORT_PATH & "rr.rpt", "{rrheader.rrnumber} = '" & txtrrNumber.Text & "'", OSMS_DataConn, 1
    'PrintSQLReport rptMRR, OSMS_Report_Path & "RR_Format.rpt", "{rrheader.rrnumber} = '" & txtrrNumber.Text & "'", OSMS_DataConn, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdAddDetail_Click()
    fmeRRDetails.Caption = "Add A Record"
    AddorEdit = "ADD"
    ShowHidePictureBox picRRDetails.hwnd, True, Me
    fmeRRDetails.Enabled = True
    grid.Enabled = True
    cmdRRDdelete.Visible = False
    On Error Resume Next
    cboSupplyCode.SetFocus
    InitMemVarsRRD
End Sub

Private Sub cmdRRDCancel_Click()
    ShowHidePictureBox picRRDetails.hwnd, False, Me
    StoreMemVars
End Sub

Private Sub cmdRRDdelete_Click()
    If MsgBoxXP("Are you sure you want to delete this record?", "Delete Current Record", XP_YesNo, msg_Question) = True Then
        gconDMIS.Execute "delete from OSMS_RRDETAILS  where id = " & labDetID.Caption
        Dim kim                                        As Integer
        Dim rsrrDETAILS                                As ADODB.Recordset
        Set rsrrDETAILS = New ADODB.Recordset
        rsrrDETAILS.Open "select * from OSMS_RRDETAILS  where rrNumber = " & N2Str2Null(txtrrNumber.Text) & " order by Item_No asc", gconDMIS
        If Not rsrrDETAILS.EOF And Not rsrrDETAILS.BOF Then
            rsrrDETAILS.MoveFirst
            kim = 0
            Do While Not rsrrDETAILS.EOF
                kim = kim + 1
                gconDMIS.Execute "update OSMS_RRDETAILS  set Item_No = '" & Format(kim, "0000") & "' where id = " & rsrrDETAILS!ID
                rsrrDETAILS.MoveNext
            Loop
        End If
    End If
    cmdRRDCancel.Value = True
End Sub

Private Sub cmdEditDetail_Click()
    grid.Enabled = True
    grid.Col = 6
    If grid.Text <> "" Then
        labDetID.Caption = grid.Text
        AddorEdit = "EDIT"
        RRDStoreMemVars
        fmeRRDetails.Caption = "Edit A Record"
        fmeRRDetails.Enabled = True
        cmdRRDdelete.Visible = True
        PrevItemNo = txtitemNo.Text
        ShowHidePictureBox picRRDetails.hwnd, True, Me
        On Error Resume Next
        cboSupply.SetFocus
    End If
End Sub

Private Sub cmdrrdSave_Click()
    Dim CheckSupply                                    As ADODB.Recordset
    Dim mysql                                          As String

    On Error GoTo ErrorHandler

    Set CheckSupply = New Recordset
    CheckSupply.Open "Select Supply_Description from OSMS_SUPPLY where Supply_Description = '" & cboSupply.Text & "'", gconDMIS
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
        mysql = "Insert into OSMS_rrDetails" & _
                "(rrNumber,item_no, rrquantity, rrunit,Cost, Supply_code) values (" & N2Str2Null(txtrrNumber.Text) & "," & N2Str2Null(txtitemNo.Text) & ", " & NumericVal(txtRRDQuantity) & ", " & N2Str2Null(SETCBOUNIT(cboUnit.Text)) & ", " & NumericVal(txtCost.Text) & ", " & N2Str2Null(cboSupplyCode.Text) & ")"
        gconDMIS.Execute mysql
    Else
        gconDMIS.Execute "update OSMS_RRDETAILS  set " & _
                         "item_no = " & N2Str2Null(txtitemNo.Text) & "," & _
                         "rrquantity = " & NumericVal(txtRRDQuantity.Text) & "," & _
                         "rrunit = " & N2Str2Null(SETCBOUNIT(cboUnit.Text)) & "," & _
                         "Cost = " & NumericVal(txtCost.Text) & "," & _
                         "Supply_Code = " & N2Str2Null(cboSupplyCode.Text) & _
                       " where id = " & labDetID.Caption
    End If
    gconDMIS.Execute "update OSMS_SUPPLY set LASTRRDATE = " & N2Date2Null(txtrrDate.Text) & ", ONHAND = ONHAND + " & NumericVal(txtRRDQuantity.Text) & ", Lastm_Cost = " & VLast_Cost & ", COST = " & NumericVal(txtCost.Text) & " where Supply_Code =  " & N2Str2Null(cboSupplyCode.Text)
    cmdRRDCancel.Value = True
    If AddorEdit = "ADD" Then cmdAddDetail_Click
    Exit Sub

ErrorHandler:
    '   MsgBoxXP "Error" & Err.Number & vbCrLf & "Description: " & Err.Description, "Error", XP_OKOnly, msg_Critical
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    Dim mysql                                          As String
    On Error GoTo ErrorHandler
    If AddorEdit = "ADD" Then
        rsrrHEADER.Find "rrNumber = '" & txtrrNumber.Text & "'"
        If Not rsrrHEADER.EOF Then
            MsgBoxXP "MRR Number already exists!", "Input New MRR Number", XP_OKOnly, msg_Exclamation
            On Error Resume Next
            txtrrNumber.SetFocus
            Exit Sub
        End If
        mysql = "INSERT INTO OSMS_RRHEADER " & _
                "(rrNumber, rrDate, INV_No,INV_Date,PO_No, PO_Date, Supplier_Code,ReceivedBy_Code,Purpose) " & _
                "values (" & N2Str2Null(txtrrNumber.Text) & "," & N2Date2Null(txtrrDate.Text) & "," & N2Str2Null(txtInv_No.Text) & "," & N2Str2Null(txtInv_Date.Text) & "," & N2Str2Null(txtPONo.Text) & "," & N2Date2Null(txtPODate.Text) & "," & N2Str2Null(SETCBOSUPPLIER(cboSupplier.Text)) & ", " & N2Str2Null(SETCBOReceivedby(cboReceivedBy.Text)) & ", " & N2Str2Null(txtRemarks.Text) & ")"
        gconDMIS.Execute mysql
        fmeRRHeader.Caption = "Materials Received Report"
    Else
        gconDMIS.Execute "UPDATE OSMS_RRHEADER  set " & _
                         "rrNumber = " & N2Str2Null(txtrrNumber.Text) & "," & _
                         "rrDate = " & N2Date2Null(txtrrDate.Text) & "," & _
                         "INV_No = " & N2Str2Null(txtInv_No.Text) & "," & _
                         "INV_Date = " & N2Str2Null(txtInv_Date.Text) & "," & _
                         "PO_No = " & N2Str2Null(txtPONo.Text) & "," & _
                         "PO_Date = " & N2Date2Null(txtPODate.Text) & "," & _
                         "Supplier_Code = " & N2Str2Null(SETCBOSUPPLIER(cboSupplier.Text)) & "," & _
                         "Purpose = " & N2Str2Null(txtRemarks.Text) & "," & _
                         "ReceivedBy_Code = " & N2Str2Null(SETCBOReceivedby(cboReceivedBy.Text)) & _
                       " where rrNumber = " & N2Str2Null(PrevRRNum)
        fmeRRHeader.Caption = "Materials Received Report"
    End If
    rsRefresh
    rsrrHEADER.Find "rrnumber = " & N2Str2Null(txtrrNumber.Text)
    cmdCancel.Value = True
    If AddorEdit = "ADD" Then cmdAddDetail_Click
    Exit Sub

ErrorHandler:
    MsgBoxXP "Error" & Err.Number & vbCrLf & "Description: " & Err.Description, "Error", XP_OKOnly, msg_Critical
End Sub

'Upating Code       : AXP-0716200719:12
'Upating Code       : AXP-0716200719:27
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Errorcode:

    Select Case KeyCode
        Case vbKeyF2
            If picAdds.Visible = True Then
                cmdAddDetail.Value = True
            End If
        Case vbKeyF3
            If picAdds.Visible = True Then
                grid_DblClick
            End If
        Case vbKeyEscape
            cmdRRDCancel.Value = True
        Case Else
            MoveKeyPress KeyCode
    End Select
Exit Sub
Errorcode:
ShowVBError
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1

    fmeRRHeader.Enabled = False
    initMemvars
    InitMemVarsRRD
    rsRefresh



    txtSearch.Text = ""
    If Not rsrrHEADER.EOF And Not rsrrHEADER.BOF Then rsrrHEADER.MoveLast
    cmdRRDCancel.Value = True
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    Set rsrrHEADER = New ADODB.Recordset
    rsrrHEADER.Open "select * from OSMS_RRHEADER  order by rrnumber asc", gconDMIS
End Sub

Sub StoreMemVars()
    If Not rsrrHEADER.EOF And Not rsrrHEADER.BOF Then
        txtrrNumber.Text = Null2String(rsrrHEADER!rrnumber)
        txtrrDate.Text = Format(Null2Date(rsrrHEADER!RRDATE), "DD-MMM-YY")
        txtInv_No.Text = Null2String(rsrrHEADER!inv_no)
        txtInv_Date.Text = Null2String(rsrrHEADER!inv_date)
        txtPONo.Text = Null2String(rsrrHEADER!PO_No)
        txtPODate.Text = Format(Null2Date(rsrrHEADER!PO_Date), "DD-MMM-YY")
        cboReceivedBy.Text = SETCBOReceivedby2(Null2String(rsrrHEADER!Receivedby_Code))
        cboSupplier.Text = SETCBOSUPPLIER2(Null2String(rsrrHEADER!Supplier_code))
        txtRemarks.Text = Null2String(rsrrHEADER!PURPOSE)
        FillGrid

    Else
        MsgBoxXP "Record is Empty!", "No Record", XP_OKOnly, msg_Information
        cmdAdd.Value = True
    End If
End Sub

Sub RRDStoreMemVars()
    Dim FieldID                                        As Long
    grid.Col = 6
    If grid.Text <> "" Then
        FieldID = grid.Text
        Set rsrrDETAILS = New ADODB.Recordset
        rsrrDETAILS.Open "select * from OSMS_RRDETAILS  where ID  = " & FieldID, gconDMIS
        If Not rsrrDETAILS.EOF And Not rsrrDETAILS.BOF Then
            labDetID.Caption = rsrrDETAILS!ID
            txtitemNo.Text = Null2String(rsrrDETAILS!item_no)
            txtRRDQuantity.Text = NumericVal(rsrrDETAILS!rrQUANTITY)
            cboSupplyCode.Text = Null2String(rsrrDETAILS!Supply_Code)
            cboSupply.Text = SETCBOSUPPLY2(Null2String(rsrrDETAILS!Supply_Code))
            cboUnit.Text = SETCBOUNIT2(Null2String(rsrrDETAILS!rrunit))
            txtCost.Text = NumericVal(rsrrDETAILS!Cost)
            VLast_Cost = NumericVal(rsrrDETAILS!Cost)

        End If
    End If
End Sub

Sub FillGrid()
    Set rsrrDETAILS = New ADODB.Recordset
    txtTotalAmount.Text = Format(0, "###,###,##0.00")
    rsrrDETAILS.Open "select * from OSMS_RRDETAILS  where rrNumber = " & N2Str2Null(txtrrNumber.Text) & " order by item_no asc", gconDMIS
    If Not rsrrDETAILS.EOF And Not rsrrDETAILS.BOF Then
        rsrrDETAILS.MoveFirst
        cleargrid grid
        grid.ColWidth(6) = 1
        TotalMRRAMount = 0
        Do While Not rsrrDETAILS.EOF
            grid.AddItem Null2String(rsrrDETAILS!item_no) & Chr(9) & _
                                 Null2String(rsrrDETAILS!Supply_Code) & Chr(9) & _
                                 NumericVal(rsrrDETAILS!rrQUANTITY) & Chr(9) & _
                                 Null2String(rsrrDETAILS!rrunit) & Chr(9) & _
                                 Format(NumericVal(rsrrDETAILS!Cost), "###,###,##0.00") & Chr(9) & _
                                 Format(NumericVal(rsrrDETAILS!rrQUANTITY) * NumericVal(rsrrDETAILS!Cost), "###,###,##0.00") & Chr(9) & _
                                 rsrrDETAILS!ID
            TotalMRRAMount = TotalMRRAMount + (NumericVal(rsrrDETAILS!rrQUANTITY) * NumericVal(rsrrDETAILS!Cost))
            rsrrDETAILS.MoveNext
        Loop
        txtTotalAmount.Text = Format(TotalMRRAMount, "###,###,##0.00")
        If grid.Rows > 2 Then grid.RemoveItem 1
    Else
        cleargrid grid
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labF2.FontUnderline = False: labF2.FontBold = False
    labF3.FontUnderline = False: labF3.FontBold = False
End Sub

Private Sub grid_DblClick()

    If grid.Text = "No Entry" Then
        cmdAddDetail.Value = True

    Else
        cmdEditDetail.Value = True

    End If


End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: grid_DblClick
End Sub

Private Sub labF2_Click()
    If picSaves.Visible = True Then cmdAddDetail.Value = True
End Sub

Private Sub labF2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labF2.FontUnderline = True
    labF2.FontBold = True
End Sub

Private Sub labF3_Click()
    If picSaves.Visible = True Then cmdEditDetail.Value = True
End Sub

Private Sub labF3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labF3.FontUnderline = True
    labF3.FontBold = True
End Sub

Private Sub txtCost_Change()
    ComputeAmount
End Sub

Private Sub txtCost_GotFocus()
    If Val(txtCost.Text) <= 0 Then txtCost.Text = ""
End Sub

Private Sub txtCost_LostFocus()
    ComputeAmount
    If Val(txtCost.Text) <= 0 Then txtCost.Text = "0.00"
End Sub

Private Sub txtrrDate_GotFocus()
    txtrrDate.Text = Format(txtrrDate.Text, "MM/DD/YYYY")
End Sub

Private Sub txtrrDate_LostFocus()
    txtrrDate.Text = Format(txtrrDate.Text, "DD-MMM-YY")
End Sub

Private Sub txtPODate_GotFocus()
    txtrrDate.Text = Format(txtrrDate.Text, "MM/DD/YYYY")
End Sub

Private Sub txtPODate_LostFocus()
    txtrrDate.Text = Format(txtrrDate.Text, "DD-MMM-YY")
End Sub

Private Sub txtRRDQuantity_Change()
    ComputeAmount
End Sub

Private Sub txtRRDQuantity_GotFocus()
    If txtRRDQuantity.Text = 1 Then
        txtRRDQuantity.Text = ""
    End If
End Sub

Private Sub txtRRDQuantity_LostFocus()
    ComputeAmount
    If txtRRDQuantity.Text = "" Then
        txtRRDQuantity.Text = 1
    End If
End Sub

Sub ComputeAmount()
    txtAmount.Text = Format(Val(txtRRDQuantity.Text) * Val(txtCost.Text), "###,###,##0.00")
End Sub



Private Sub lstReceiving_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsrrHEADER.Bookmark = rsFind(rsrrHEADER.Clone, "RRNumber", lstReceiving.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstReceiving_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstReceiving
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

Private Sub lstReceiving_DblClick()
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
    Dim rsrrHEADER2                                    As ADODB.Recordset
    lstReceiving.Sorted = False: lstReceiving.ListItems.Clear
    lstReceiving.Enabled = False
    Set rsrrHEADER2 = New ADODB.Recordset
    Set rsrrHEADER2 = gconDMIS.Execute("select RRDate,RRNumber from OSMS_RRHEADER  order by RRDate asc")
    If Not (rsrrHEADER2.EOF And rsrrHEADER2.BOF) Then
        Listview_Loadval Me.lstReceiving.ListItems, rsrrHEADER2
        lstReceiving.Refresh
    End If
    lstReceiving.Enabled = True
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsrrHEADER2                                    As ADODB.Recordset
    lstReceiving.Sorted = False: lstReceiving.ListItems.Clear
    lstReceiving.Enabled = False
    Set rsrrHEADER2 = New ADODB.Recordset
    Set rsrrHEADER2 = gconDMIS.Execute("select RRDate,RRNumber from OSMS_RRHEADER  where RRDate like'" & XXX & "%' order by RRDate asc")
    If Not (rsrrHEADER2.EOF And rsrrHEADER2.BOF) Then
        Listview_Loadval Me.lstReceiving.ListItems, rsrrHEADER2
        lstReceiving.Refresh
        lstReceiving.Enabled = True
    End If
   
End Sub

Sub FillGrid1()
    Dim rsrrHEADER2                                    As ADODB.Recordset
    lstReceiving.Sorted = False: lstReceiving.ListItems.Clear
    lstReceiving.Enabled = False
    Set rsrrHEADER2 = New ADODB.Recordset
    Set rsrrHEADER2 = gconDMIS.Execute("select RRNumber,RRNumber from OSMS_RRHEADER  order by RRNumber asc")
    If Not (rsrrHEADER2.EOF And rsrrHEADER2.BOF) Then
        Listview_Loadval Me.lstReceiving.ListItems, rsrrHEADER2
        lstReceiving.Refresh
        lstReceiving.Enabled = True
    End If
    
End Sub

Sub FillSearchGrid1(XXX As String)
    Dim rsrrHEADER2                                    As ADODB.Recordset
    lstReceiving.Enabled = False
    lstReceiving.Sorted = False: lstReceiving.ListItems.Clear
    Set rsrrHEADER2 = New ADODB.Recordset
    Set rsrrHEADER2 = gconDMIS.Execute("select RRNumber,RRNumber from OSMS_RRHEADER  where RRNumber like'" & ReplaceQuote(XXX) & "%' order by RRNumber asc")
    If Not (rsrrHEADER2.EOF And rsrrHEADER2.BOF) Then
        Listview_Loadval Me.lstReceiving.ListItems, rsrrHEADER2
        lstReceiving.Refresh
        lstReceiving.Enabled = True
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





