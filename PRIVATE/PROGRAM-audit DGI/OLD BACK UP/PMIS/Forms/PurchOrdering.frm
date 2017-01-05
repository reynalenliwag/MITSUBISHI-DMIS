VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "xpbutton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISPurchProcessing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Processing"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11760
   ForeColor       =   &H00DEDFDE&
   Icon            =   "PurchOrdering.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   11760
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2190
      ScaleHeight     =   255
      ScaleWidth      =   9465
      TabIndex        =   70
      Top             =   5040
      Width           =   9495
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
         Left            =   7110
         TabIndex        =   75
         Top             =   30
         Width           =   2445
      End
      Begin VB.Label Label24 
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
         Left            =   5070
         TabIndex        =   74
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label25 
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
         Left            =   3360
         TabIndex        =   73
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label26 
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
         Left            =   1740
         TabIndex        =   72
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label27 
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
         Left            =   120
         TabIndex        =   71
         Top             =   30
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2985
      Left            =   2190
      TabIndex        =   11
      Top             =   0
      Width           =   9495
      Begin VB.TextBox txtDS1 
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
         Left            =   5550
         MaxLength       =   3
         TabIndex        =   8
         ToolTipText     =   "Type the percentage of the amount to be added. Do not include % sign (e.g. 10, 15)"
         Top             =   1020
         Width           =   465
      End
      Begin VB.TextBox txtPPDate 
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
         Left            =   6060
         MaxLength       =   10
         TabIndex        =   50
         ToolTipText     =   "Type purchase order date in mm/dd/yyyy format (e.g. 7/5/2004)"
         Top             =   180
         Width           =   1365
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
         Height          =   945
         Left            =   5550
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "PurchOrdering.frx":08CA
         ToolTipText     =   "Type your message or your remarks."
         Top             =   1980
         Width           =   3855
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   90
         ScaleHeight     =   795
         ScaleWidth      =   5355
         TabIndex        =   40
         Top             =   2130
         Width           =   5355
         Begin VB.TextBox txtSHP_addrs 
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
            Height          =   795
            Left            =   0
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   0
            Width           =   5265
         End
      End
      Begin VB.TextBox txtPPNo 
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
         Left            =   4110
         MaxLength       =   7
         TabIndex        =   1
         ToolTipText     =   "Type the purchase order number (e.g. 002775)"
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txtShipTo 
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
         Left            =   2010
         MaxLength       =   40
         TabIndex        =   5
         ToolTipText     =   "Type the name of addressee (e.g. CALEB MOTOR CORPORATION)"
         Top             =   1770
         Width           =   3375
      End
      Begin VB.TextBox txtDealerCode 
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
         Left            =   960
         MaxLength       =   6
         TabIndex        =   4
         ToolTipText     =   "Type the place where the order should be delivered (e.g. PCMC0)"
         Top             =   1770
         Width           =   1005
      End
      Begin VB.TextBox txtSupCode 
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
         Left            =   960
         MaxLength       =   6
         TabIndex        =   0
         ToolTipText     =   "Type supplier code (e.g. 00001, 00002)"
         Top             =   180
         Width           =   1365
      End
      Begin VB.ComboBox cboSupName 
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
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "cboRecvd_Desc"
         ToolTipText     =   "Select supplier name from the list."
         Top             =   630
         Width           =   5025
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
         Left            =   6420
         TabIndex        =   9
         ToolTipText     =   "Input the type of the additional amount (e.g. VAT)"
         Top             =   1020
         Width           =   1425
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1215
         Left            =   7920
         ScaleHeight     =   1215
         ScaleWidth      =   1545
         TabIndex        =   18
         Top             =   600
         Width           =   1545
         Begin VB.TextBox txtPP_Amount 
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
            Left            =   120
            MaxLength       =   15
            TabIndex        =   56
            Top             =   30
            Width           =   1395
         End
         Begin VB.TextBox txtNetPPAmt 
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
            Left            =   120
            MaxLength       =   15
            TabIndex        =   55
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
            Left            =   30
            MaxLength       =   15
            TabIndex        =   54
            Top             =   420
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   90
         ScaleHeight     =   765
         ScaleWidth      =   5355
         TabIndex        =   17
         Top             =   990
         Width           =   5355
         Begin VB.TextBox txtSup_Addrs 
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
            Height          =   765
            Left            =   0
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   0
            Width           =   5265
         End
      End
      Begin VB.Label Label2 
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
         Height          =   225
         Index           =   6
         Left            =   5160
         TabIndex        =   69
         Top             =   660
         Width           =   135
      End
      Begin VB.Label Label2 
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
         Height          =   225
         Index           =   5
         Left            =   7440
         TabIndex        =   68
         Top             =   180
         Width           =   135
      End
      Begin VB.Label Label2 
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
         Height          =   225
         Index           =   2
         Left            =   2340
         TabIndex        =   58
         Top             =   180
         Width           =   135
      End
      Begin VB.Label Label2 
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
         Height          =   225
         Index           =   0
         Left            =   5220
         TabIndex        =   57
         Top             =   210
         Width           =   135
      End
      Begin VB.Label Label17 
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
         Left            =   6060
         TabIndex        =   51
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label labPPsted 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
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
         Left            =   7650
         TabIndex        =   48
         Top             =   210
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Order Number"
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
         Left            =   2700
         TabIndex        =   41
         Top             =   210
         Width           =   1395
      End
      Begin VB.Label Label8 
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
         Height          =   285
         Left            =   5580
         TabIndex        =   39
         Top             =   1710
         Width           =   1965
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To"
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
         Index           =   1
         Left            =   90
         TabIndex        =   38
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label2 
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
         Height          =   315
         Index           =   1
         Left            =   5550
         TabIndex        =   16
         Top             =   210
         Width           =   1965
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         TabIndex        =   15
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
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
         Left            =   6750
         TabIndex        =   14
         Top             =   690
         Width           =   1245
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
         Left            =   6750
         TabIndex        =   13
         Top             =   1440
         Width           =   1245
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
         Left            =   3960
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.Timer Timer1 
      Left            =   9660
      Top             =   120
   End
   Begin Crystal.CrystalReport rptPurchaseOrder 
      Left            =   2400
      Top             =   5790
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame fraDetails 
      BorderStyle     =   0  'None
      Height          =   2145
      Left            =   2190
      TabIndex        =   10
      Top             =   2880
      Width           =   9495
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   1905
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   3360
         _Version        =   393216
         Cols            =   9
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
      Begin VB.Label Label2 
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
         Height          =   225
         Index           =   3
         Left            =   1770
         TabIndex        =   59
         Top             =   870
         Width           =   135
      End
   End
   Begin VB.Frame fraAddTran 
      Height          =   4755
      Left            =   4200
      TabIndex        =   19
      Top             =   150
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
         Left            =   3600
         MouseIcon       =   "PurchOrdering.frx":08E4
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":0A36
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   3750
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
         Left            =   2175
         MouseIcon       =   "PurchOrdering.frx":0D61
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":0EB3
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   3750
         Width           =   705
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
         Left            =   2887
         MouseIcon       =   "PurchOrdering.frx":1203
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":1355
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   3750
         Width           =   705
      End
      Begin VB.TextBox txtTranTotalAmt 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   35
         Top             =   3420
         Width           =   1635
      End
      Begin VB.TextBox txtUnitCost 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   34
         Top             =   3060
         Width           =   1635
      End
      Begin VB.TextBox txtTranINVAmt 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   31
         Top             =   1980
         Width           =   1635
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   30
         Top             =   1620
         Width           =   885
      End
      Begin VB.TextBox txtUnit 
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
         MaxLength       =   3
         TabIndex        =   32
         Top             =   2340
         Width           =   1635
      End
      Begin VB.TextBox txtTRemarks 
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
         MaxLength       =   50
         TabIndex        =   33
         Top             =   2700
         Width           =   2925
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
         MaxLength       =   4
         TabIndex        =   27
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox cboTranDescription 
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
         Sorted          =   -1  'True
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   1230
         Width           =   4365
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
         Left            =   1470
         Sorted          =   -1  'True
         TabIndex        =   28
         Text            =   "Combo1"
         Top             =   600
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
         Left            =   1590
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
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
         Left            =   990
         TabIndex        =   47
         Top             =   2370
         Width           =   435
      End
      Begin VB.Label Label4 
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
         Left            =   570
         TabIndex        =   46
         Top             =   2730
         Width           =   855
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
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
         Left            =   120
         TabIndex        =   20
         Top             =   3420
         Width           =   1305
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
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
         Left            =   420
         TabIndex        =   37
         Top             =   3060
         Width           =   1005
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
         Height          =   105
         Left            =   1620
         TabIndex        =   26
         Top             =   3480
         Width           =   285
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
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
         Left            =   150
         TabIndex        =   25
         Top             =   2010
         Width           =   1275
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
         Height          =   285
         Left            =   540
         TabIndex        =   24
         Top             =   1620
         Width           =   885
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
         Height          =   255
         Left            =   540
         TabIndex        =   23
         Top             =   630
         Width           =   885
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
         Left            =   540
         TabIndex        =   22
         Top             =   270
         Width           =   885
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
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6495
      Left            =   60
      TabIndex        =   62
      Top             =   0
      Width           =   2115
      Begin VB.OptionButton optPPNo 
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
         TabIndex        =   65
         Top             =   390
         Value           =   -1  'True
         Width           =   1845
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
         TabIndex        =   64
         Top             =   630
         Width           =   1875
      End
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
         TabIndex        =   63
         Text            =   "TEXT"
         Top             =   975
         Width           =   1995
      End
      Begin MSComctlLib.ListView lstPP_HD 
         Height          =   5085
         Left            =   60
         TabIndex        =   66
         Top             =   1320
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
         MouseIcon       =   "PurchOrdering.frx":1693
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
         TabIndex        =   67
         Top             =   150
         Width           =   1455
      End
   End
   Begin wizButton.cmd cmdAddTran 
      Height          =   4905
      Left            =   4125
      TabIndex        =   52
      Top             =   90
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   8652
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
      MICON           =   "PurchOrdering.frx":17F5
   End
   Begin VB.Frame fraPrintFormat 
      Caption         =   "Print Format"
      Height          =   1845
      Left            =   5415
      TabIndex        =   42
      Top             =   2640
      Width           =   2445
      Begin wizButton.cmd cmdMMPC 
         Height          =   435
         Left            =   120
         TabIndex        =   43
         Top             =   300
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   767
         TX              =   "&MMPC"
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
         MICON           =   "PurchOrdering.frx":1811
      End
      Begin wizButton.cmd cmdDMC 
         Height          =   435
         Left            =   120
         TabIndex        =   44
         Top             =   780
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   767
         TX              =   "&DIAMOND"
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
         MICON           =   "PurchOrdering.frx":182D
      End
      Begin wizButton.cmd cmdOutPurch 
         Height          =   435
         Left            =   120
         TabIndex        =   45
         Top             =   1260
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   767
         TX              =   "&OUTSIDE PURCHASE"
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
         MICON           =   "PurchOrdering.frx":1849
      End
   End
   Begin wizButton.cmd cmdPrintFormat 
      Height          =   2055
      Left            =   5355
      TabIndex        =   53
      Top             =   2580
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   3625
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
      MICON           =   "PurchOrdering.frx":1865
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   2355
      ScaleHeight     =   870
      ScaleWidth      =   9285
      TabIndex        =   79
      Top             =   5400
      Width           =   9285
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
         Left            =   2295
         MouseIcon       =   "PurchOrdering.frx":1881
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":19D3
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   15
         Width           =   765
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
         Left            =   3060
         MouseIcon       =   "PurchOrdering.frx":1D31
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":1E83
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   15
         Width           =   765
      End
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
         Left            =   8445
         MouseIcon       =   "PurchOrdering.frx":21D3
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":2325
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   15
         Width           =   765
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
         Left            =   7680
         MouseIcon       =   "PurchOrdering.frx":268B
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":27DD
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   15
         Width           =   765
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
         Left            =   4590
         MouseIcon       =   "PurchOrdering.frx":2B43
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":2C95
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   15
         Width           =   765
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
         Left            =   3825
         MouseIcon       =   "PurchOrdering.frx":2FF1
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":3143
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   15
         Width           =   765
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
         Left            =   1530
         MouseIcon       =   "PurchOrdering.frx":3456
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":35A8
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   15
         Width           =   765
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
         Left            =   765
         MouseIcon       =   "PurchOrdering.frx":38A2
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":39F4
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   15
         Width           =   765
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
         Left            =   0
         MouseIcon       =   "PurchOrdering.frx":3D4C
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":3E9E
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   15
         Width           =   765
      End
      Begin VB.CommandButton cmdCancelPP 
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
         Left            =   6900
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "PurchOrdering.frx":41FD
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":434F
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   15
         Width           =   765
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
         Left            =   6135
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "PurchOrdering.frx":4689
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":47DB
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   15
         Width           =   765
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
         Left            =   5355
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "PurchOrdering.frx":4B20
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":4C72
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Press F11 for Posting By Range"
         Top             =   15
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   10170
      ScaleHeight     =   855
      ScaleWidth      =   1470
      TabIndex        =   92
      Top             =   5400
      Width           =   1470
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
         Left            =   705
         MouseIcon       =   "PurchOrdering.frx":4F97
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":50E9
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   0
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
         Left            =   0
         MouseIcon       =   "PurchOrdering.frx":5427
         MousePointer    =   99  'Custom
         Picture         =   "PurchOrdering.frx":5579
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "- required fields"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   10230
      TabIndex        =   61
      Top             =   6330
      Width           =   1425
   End
   Begin VB.Label Label2 
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
      Height          =   225
      Index           =   4
      Left            =   10050
      TabIndex        =   60
      Top             =   6360
      Width           =   135
   End
End
Attribute VB_Name = "frmPMISPurchProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPP_HD, rsPPDaytran, rsPartMas    As ADODB.Recordset
Attribute rsPPDaytran.VB_VarUserMemId = 1073938432
Attribute rsPartMas.VB_VarUserMemId = 1073938432
Dim rsSupplier, rsProfile, rsCunter    As ADODB.Recordset
Attribute rsSupplier.VB_VarUserMemId = 1073938435
Attribute rsProfile.VB_VarUserMemId = 1073938435
Attribute rsCunter.VB_VarUserMemId = 1073938435
Dim Pcnt                               As Integer
Attribute Pcnt.VB_VarUserMemId = 1073938438
Dim AddorEdit                          As String
Attribute AddorEdit.VB_VarUserMemId = 1073938439
Dim PP_TOTUCOST, PP_TOTINVAMT, PP_TOTVAT As Double
Attribute PP_TOTUCOST.VB_VarUserMemId = 1073938440
Attribute PP_TOTINVAMT.VB_VarUserMemId = 1073938440
Attribute PP_TOTVAT.VB_VarUserMemId = 1073938440
Dim PP_T_ONORDER                       As Long
Attribute PP_T_ONORDER.VB_VarUserMemId = 1073938443
Dim kcnt                               As Integer
Attribute kcnt.VB_VarUserMemId = 1073938444
Dim PrevPPNO                           As String
Attribute PrevPPNO.VB_VarUserMemId = 1073938445

Private Sub cboSupName_Click()
    txtSupCode.Text = SetSupCode(cboSupName.Text)
End Sub

Private Sub cboSupName_LostFocus()
    txtSupCode.Text = SetSupCode(cboSupName.Text)
End Sub

Private Sub cboTranDescription_Click()
    If cboTranDescription.Text <> "" Then
        txtPartID.Text = SetPartIDDesc(cboTranDescription.Text)
        cboTranPartNo.Text = SetSTOCKNO(txtPartID.Text)
        cboTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
    End If
End Sub

Private Sub cboTranPartNo_Change()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        cboTranDescription.Text = SetSTOCKDESC2(NumericVal(txtPartID.Text))
    End If
End Sub

Private Sub cboTranPartNo_Click()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        cboTranDescription.Text = SetSTOCKDESC2(NumericVal(txtPartID.Text))
    End If
End Sub

Private Sub cmdAddTran_Click()
    If Picture1.Visible = True Then
        SendToBack
        cmdAddTran.ZOrder 0
        fraAddTran.ZOrder 0
        fraAddTran.Enabled = True
        AddorEdit = "ADD"
        InitParts
        cboTranPartNo.SetFocus
    End If
End Sub

Private Sub cmdCancelPP_Click()
    If Function_Access(LOGID, "Acess_CancelEntry") = False Then Exit Sub

    'If LOGLEVEL <> "ADM" Then
    '   MsgBox "Warning: Your account is not allowed to cancel this transaction!", vbCritical, "Error"
    '   Exit Sub
    'End If
    Dim rsPPDaytranDup, rsPartmasDup   As ADODB.Recordset
    Dim PCurOnOrder, PCurTppQty        As Integer

    If MsgBoxXP("Are you sure you want to Cancel this Transactions?", "Cancel Transactions", XP_YesNo, msg_Question) = True Then
        gconDMIS.Execute "update PMIS_PP_Hd set" & _
                       " status = 'C'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labid.Caption
        gconDMIS.Execute "update PMIS_PPDayTran set" & _
                       " status = 'C'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where tranno = " & N2Str2Null(rsPP_HD!ppno) & " and trantype = 'PP'"
        Set rsPPDaytranDup = New ADODB.Recordset
        rsPPDaytranDup.Open "select trantype,tranno,PART_ORD,tranqty from PMIS_PPDayTran where trantype = 'PP' and tranno = " & N2Str2Null(rsPP_HD!ppno), gconDMIS
        If Not rsPPDaytranDup.EOF And Not rsPPDaytranDup.BOF Then
            rsPPDaytranDup.MoveFirst
            Do While Not rsPPDaytranDup.EOF
                Set rsPartmasDup = New ADODB.Recordset
                rsPartmasDup.Open "select onorder,tppqty,STOCKNO from PMIS_STOCKMAS where STOCKNO = " & N2Str2Null(rsPPDaytranDup!PART_ORD), gconDMIS
                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                    PCurOnOrder = N2Str2Zero(rsPartmasDup!onorder) - N2Str2Zero(rsPPDaytranDup!tranqty)
                    PCurTppQty = N2Str2Zero(rsPartmasDup!tppqty) - N2Str2Zero(rsPPDaytranDup!tranqty)
                    gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                                   " onorder = " & PCurOnOrder & "," & _
                                   " tppqty = " & PCurTppQty & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where STOCKNO = " & N2Str2Null(rsPPDaytranDup!PART_ORD)
                End If
                rsPPDaytranDup.MoveNext
            Loop
        End If
        rsRefresh
        On Error Resume Next
        rsPP_HD.Find "id =" & labid.Caption
        StoreMemVars
    End If
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete") = False Then Exit Sub

End Sub

Private Sub cmdDMC_Click()
    Screen.MousePointer = 11
    rptPurchaseOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPurchaseOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "PP.RPT", "{PP_hd.PPno} = '" & txtPPNo.Text & "'", DMIS_REPORT_Connection, 1
    fraPrintFormat.Visible = False
    cmdPrintFormat.Visible = False
    cmdPrintFormat.ZOrder 0
    fraPrintFormat.ZOrder 0
    Screen.MousePointer = 0
End Sub

Private Sub cmdMMPC_Click()
    Screen.MousePointer = 11
    rptPurchaseOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPurchaseOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "PP2.RPT", "{PP_hd.PPno} = " & N2Str2Null(txtPPNo.Text), DMIS_REPORT_Connection, 1
    fraPrintFormat.Visible = False
    cmdPrintFormat.Visible = False
    cmdPrintFormat.ZOrder 0
    fraPrintFormat.ZOrder 0
    Screen.MousePointer = 0
End Sub

Private Sub cmdOutPurch_Click()
    Screen.MousePointer = 11
    rptPurchaseOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPurchaseOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "outpurch.RPT", "{PP_hd.PPno} = " & N2Str2Null(txtPPNo.Text), DMIS_REPORT_Connection, 1
    fraPrintFormat.Visible = False
    cmdPrintFormat.Visible = False
    cmdPrintFormat.ZOrder 0
    fraPrintFormat.ZOrder 0
    Screen.MousePointer = 0
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post") = False Then Exit Sub
    If Null2String(rsPP_HD!Status) = "C" Then
        MsgSpeechBox "Transaction has been Cancelled and can not be Posted"
        Exit Sub
    End If
    Set rsPPDaytran = New ADODB.Recordset
    rsPPDaytran.Open "select trantype,tranno from PMIS_PPDayTran where trantype = 'PP' and tranno = " & N2Str2Null(rsPP_HD!ppno), gconDMIS
    If rsPPDaytran.EOF And rsPPDaytran.BOF Then
        MsgSpeechBox "This Transaction has no detail and cannot be Posted!"
        Exit Sub
    End If
    If MsgQuestionBox("Are you sure you want to Post this Transactions?", "Post Transactions") = True Then
        gconDMIS.Execute "update PMIS_PP_Hd set" & _
                       " status = 'P'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labid.Caption
        gconDMIS.Execute "update PMIS_PPDayTran set" & _
                       " status = 'P'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where trantype = 'PP' and tranno = " & N2Str2Null(rsPP_HD!ppno)
        LogAudit "P", "Purchase Processing", txtPPNo
        rsRefresh
        rsPP_HD.Find "id =" & labid.Caption
        StoreMemVars
    End If
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print") = False Then Exit Sub

    fraPrintFormat.Visible = True
    cmdPrintFormat.Visible = True
    cmdPrintFormat.ZOrder 0
    fraPrintFormat.ZOrder 0
End Sub

Private Sub cmdTranCancel_Click()
    SendToBack
    StoreMemVars
End Sub

Private Sub cmdTranDelete_Click()
    If MsgQuestionBox("Delete This Order, Are you Sure?", "Delete Parts Entry") = True Then
        gconDMIS.Execute "delete from PMIS_PPDayTran where id = " & labDetID.Caption
    End If
    FillDetails
    PP_TOTVAT = PP_TOTINVAMT - PP_TOTUCOST
    gconDMIS.Execute "update PMIS_PP_Hd set" & _
                   " PP_amount = " & PP_TOTUCOST & "," & _
                   " netPPamt = " & PP_TOTINVAMT & "," & _
                   " ds_desc1 = '" & "VAT" & "'," & _
                   " ds_amt1 = " & PP_TOTVAT & _
                   " where id = " & labid.Caption
    Dim cnt                            As Integer
    Dim rsPPDaytranDup                 As ADODB.Recordset
    Set rsPPDaytranDup = New ADODB.Recordset
    rsPPDaytranDup.Open "select id,itemno from PMIS_PPDayTran where trantype = 'PP' and tranno = " & N2Str2Null(rsPP_HD!ppno) & " order by itemno asc", gconDMIS
    If Not rsPPDaytranDup.EOF And Not rsPPDaytranDup.BOF Then
        rsPPDaytranDup.MoveFirst
        cnt = 0
        Do While Not rsPPDaytranDup.EOF
            cnt = cnt + 1
            gconDMIS.Execute "update PMIS_PPDayTran set itemno = '" & Format(cnt, "0000") & "' where id = " & rsPPDaytranDup!ID
            rsPPDaytranDup.MoveNext
        Loop
    End If
    rsRefresh
    On Error Resume Next
    rsPP_HD.Find "id = " & labid.Caption
    cmdTranCancel.Value = True
End Sub

Private Sub cmdTranSave_Click()
    On Error GoTo ErrorCode

    If cboTranPartNo.Text = "" Then
        MsgSpeechBox "Part Number must have a value"
        On Error Resume Next
        cboTranPartNo.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Dim rsPPDaytranClone           As ADODB.Recordset
        Set rsPPDaytranClone = New ADODB.Recordset
        rsPPDaytranClone.Open "select trantype,tranno,itemno,PART_ORD from PMIS_PPDayTran where PART_ORD = '" & cboTranPartNo.Text & "' and trantype = 'PP' and tranno =" & N2Str2Null(rsPP_HD!ppno) & " order by itemno asc", gconDMIS
        If Not rsPPDaytranClone.EOF And Not rsPPDaytranClone.BOF Then
            MsgSpeechBox "Part Number already used in this transaction"
            On Error Resume Next
            cboTranPartNo.SetFocus
            Exit Sub
        End If
    End If

    Dim PPTRANDATE, PPTRANNO, PPTRANTYPE As String
    Dim PPITEMNO, PPPART_ORD, PPPART_SUP As String
    Dim PPTRANQTY                      As Integer
    Dim PPUNIT, PPTREMARKS             As String
    Dim PPTRANUCOST, PPTRANINVAMT      As Double
    Dim PPSTATUS                       As String

    PPTRANDATE = N2Date2Null(txtPPDate.Text)
    PPTRANTYPE = "'PP'"
    PPTRANNO = N2Str2Null(txtPPNo.Text)
    PPITEMNO = N2Str2Null(Format(txtTranItemNo.Text, "0000"))
    PPPART_ORD = N2Str2Null(cboTranPartNo.Text)
    PPPART_SUP = N2Str2Null(Mid(cboTranDescription.Text, 1, 50))
    PPTRANQTY = NumericVal(txtTranQty.Text)
    PPTRANINVAMT = NumericVal(txtTranINVAmt.Text)
    PPTRANUCOST = NumericVal(txtUnitCost.Text)
    PPUNIT = N2Str2Null(txtUnit.Text)
    PPTREMARKS = N2Str2Null(txtTRemarks.Text)
    PPSTATUS = "'N'"

    Dim TPO_T_ONORDER                  As Long
    Dim rsPartMasClone                 As ADODB.Recordset

    Set rsPartMasClone = New ADODB.Recordset
    rsPartMasClone.Open "select STOCKNO,onorder from PMIS_STOCKMAS where STOCKNO = '" & cboTranPartNo.Text & "'", gconDMIS
    If Not rsPartMasClone.EOF And Not rsPartMasClone.BOF Then
        TPO_T_ONORDER = N2Str2Zero(rsPartMasClone!onorder) + PPTRANQTY
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " onorder = " & TPO_T_ONORDER & _
                       " where STOCKNO = " & PPPART_ORD
    End If

    Screen.MousePointer = 11
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "insert into PMIS_PPDayTran " & _
                         "(trandate,trantype,tranno,itemno,PART_ORD,PART_SUP,tranqty,tranucost,traninvamt,unit,tremarks,lastupdate,usercode,status)" & _
                       " values (" & PPTRANDATE & ", " & PPTRANTYPE & ", " & PPTRANNO & "," & _
                       " " & PPITEMNO & "," & PPPART_ORD & "," & _
                       " " & PPPART_SUP & ", " & PPTRANQTY & "," & _
                       " " & PPTRANUCOST & ", " & PPTRANINVAMT & ", " & PPUNIT & ", " & PPTREMARKS & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & PPSTATUS & ")"
    Else
        gconDMIS.Execute "update PMIS_PPDayTran set" & _
                       " trandate = " & PPTRANDATE & "," & _
                       " trantype = " & PPTRANTYPE & "," & _
                       " tranno = " & PPTRANNO & "," & _
                       " itemno = " & PPITEMNO & "," & _
                       " PART_ORD = " & PPPART_ORD & "," & _
                       " PART_SUP = " & PPPART_SUP & "," & _
                       " tranqty = " & PPTRANQTY & "," & _
                       " tranucost = " & PPTRANUCOST & "," & _
                       " traninvamt = " & PPTRANINVAMT & "," & _
                       " unit = " & PPUNIT & "," & _
                       " tremarks = " & PPTREMARKS & "," & _
                       " lastupdate = '" & LOGDATE & "'," & _
                       " status = " & PPSTATUS & "," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "" & _
                       " where id = " & labDetID.Caption
    End If
    cleargrid grdDetails
    FillDetails
    PP_TOTVAT = PP_TOTINVAMT - PP_TOTUCOST
    gconDMIS.Execute "update PMIS_PP_Hd set" & _
                   " PP_amount = " & PP_TOTUCOST & "," & _
                   " netPPamt = " & PP_TOTINVAMT & "," & _
                   " ds_desc1 = '" & "VAT" & "'," & _
                   " ds_amt1 = " & PP_TOTVAT & _
                   " where id = " & labid.Caption
    rsRefresh
    On Error Resume Next
    rsPP_HD.Find "id = " & labid.Caption
    StoreMemVars
    If AddorEdit = "ADD" Then cmdAddTran_Click Else cmdTranCancel.Value = True
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
    Exit Sub
End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UnPost") = False Then Exit Sub

    'If LOGLEVEL <> "ADM" Then
    '   MsgBox "Warning: Your account is not allowed to unpost this transaction!", vbCritical, "Error"
    '   Exit Sub
    'End If
    If MsgQuestionBox("Are you sure you want to Unpost this Transactions?", "Unpost Transactions") = True Then
        gconDMIS.Execute "update PMIS_PP_Hd set" & _
                       " status = 'N'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labid.Caption
        gconDMIS.Execute "update PMIS_PPDayTran set" & _
                       " status = 'N'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where trantype = 'PP' and tranno = " & N2Str2Null(rsPP_HD!ppno)
        LogAudit "U", "Purchase Processing", txtPPNo
        rsRefresh
        rsPP_HD.Find "id =" & labid.Caption
        StoreMemVars
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add") = False Then Exit Sub
    AddorEdit = "ADD"
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    InitMemvars
    On Error Resume Next
    txtPPNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit") = False Then Exit Sub
    AddorEdit = "EDIT"
    PrevPPNO = Format(txtPPNo.Text, "0000000")
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    textSearch.SetFocus
    'Picture6.Visible = False
End Sub

Private Sub cmdFirst_Click()
    rsPP_HD.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    rsPP_HD.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    rsPP_HD.MoveNext
    If rsPP_HD.EOF Then
        rsPP_HD.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsPP_HD.MovePrevious
    If rsPP_HD.BOF Then
        rsPP_HD.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    Dim NewPPCunter                    As String
    Dim rsPP_HDDup                     As ADODB.Recordset

    If txtSupCode.Text = "" Then
        MsgSpeechBox "Supplier Code must not be empty!"
        On Error Resume Next
        txtSupCode.SetFocus
        Exit Sub
    End If
    If txtPPDate.Text = "" Or IsDate(txtPPDate.Text) = False Then
        MsgSpeechBox "Invalid Date!"
        On Error Resume Next
        txtPPDate.SetFocus
        Exit Sub
    End If
    If IsNull(txtPPNo.Text) = True Or Len(txtPPNo.Text) = 0 Then
        MsgSpeechBox "Purchase Processing Number must not be empty"
        On Error Resume Next
        txtPPNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsPP_HDDup = New ADODB.Recordset
            rsPP_HDDup.Open "select ppno from PMIS_PP_Hd where PPno = '" & txtPPNo.Text & "' and ordertype = '" & ORDERTYPE & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsPP_HDDup.EOF And Not rsPP_HDDup.BOF Then
                MsgSpeechBox "Purchase Processing Number already exist!"
                On Error Resume Next
                txtPPNo.SetFocus
                Exit Sub
            End If
        End If
    End If

    If Len(txtPPNo.Text) = 7 Then
        NewPPCunter = Mid(txtPPNo.Text, 1, Len(txtPPNo.Text) - 3) & Format(NumericVal(Right(txtPPNo.Text, 3)) + 1, "000")
    Else
        NewPPCunter = N2Str2Null(txtPPNo.Text)
    End If

    Dim VTXTPPNo, VTXTPPDate, VTXTSupCode, VcboSupName As String
    Dim VTXTSup_Addrs, VTXTDealerCode, VTXTShipTo, VTXTSHP_Addrs As String
    Dim VTXTPP_Amount, VTXTDS1         As Double
    Dim VTXTDS_Desc1                   As String
    Dim VTXTDS_Amt1, VTXTNetPPAmt      As Double
    Dim VTXTRemarks                    As String
    Dim VORDERTYPE                     As String

    VTXTPPNo = N2Str2Null(txtPPNo.Text)
    VTXTPPDate = N2Date2Null(txtPPDate.Text)
    VTXTSupCode = N2Str2Null(txtSupCode.Text)
    VcboSupName = N2Str2Null(Trim(cboSupName.Text))
    VTXTSup_Addrs = (N2Str2Null(Trim(txtSup_Addrs.Text)))
    VTXTDealerCode = N2Str2Null(txtDealerCode.Text)
    VTXTShipTo = N2Str2Null(txtShipTo.Text)
    VTXTSHP_Addrs = N2Str2Null(Trim(txtSHP_addrs.Text))
    VTXTPP_Amount = NumericVal(txtPP_Amount.Text)
    VTXTDS1 = NumericVal(txtDS1.Text)
    VTXTDS_Desc1 = N2Str2Null(txtDS_Desc1.Text)
    VTXTDS_Amt1 = NumericVal(txtDS_Amt1.Text)
    VTXTNetPPAmt = NumericVal(txtNetPPAmt.Text)
    VORDERTYPE = "'" & ORDERTYPE & "'"
    If txtRemarks.Text = "Pls Type Your Message Here!" Then
        VTXTRemarks = "NULL"
    Else
        VTXTRemarks = N2Str2Null(Trim(txtRemarks.Text))
    End If

    If AddorEdit = "ADD" Then
        Set rsPP_HDDup = New ADODB.Recordset
        rsPP_HDDup.Open "select id from PMIS_PP_Hd  order by id desc", gconDMIS
        If Not rsPP_HDDup.EOF And Not rsPP_HDDup.BOF Then
            rsPP_HDDup.MoveFirst
            labid.Caption = NumericVal(rsPP_HDDup!ID) + 1
        End If
        gconDMIS.Execute "Insert into PMIS_PP_Hd" & _
                       " (ordertype,PPno,PPdate,supcode,supname,sup_addrs,dealercode,PP_amount,ds1,ds_desc1,ds_amt1,netPPamt,usercode,lastupdate,remarks)" & _
                       " values (" & VORDERTYPE & ", " & VTXTPPNo & ", " & VTXTPPDate & ", " & _
                       " " & VTXTSupCode & ", " & VcboSupName & _
                         ", " & VTXTSup_Addrs & ", " & VTXTDealerCode & _
                         ", " & VTXTPP_Amount & _
                         ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                         ", " & VTXTNetPPAmt & ", " & N2Str2Null(LOGCODE) & ", '" & LOGDATE & "'," & VTXTRemarks & ")"
        LogAudit "A", "Purchase Processing"
    Else
        gconDMIS.Execute "update PMIS_PP_Hd set" & _
                       " ordertype = " & VORDERTYPE & "," & _
                       " PPno = " & VTXTPPNo & "," & _
                       " PPdate = " & VTXTPPDate & "," & _
                       " supcode = " & VTXTSupCode & "," & _
                       " supname = " & VcboSupName & "," & _
                       " sup_addrs = " & VTXTSup_Addrs & "," & _
                       " dealercode = " & VTXTDealerCode & "," & _
                       " PP_amount = " & VTXTPP_Amount & "," & _
                       " ds1 = " & VTXTDS1 & "," & _
                       " ds_desc1 = " & VTXTDS_Desc1 & "," & _
                       " ds_amt1 = " & VTXTDS_Amt1 & "," & _
                       " netPPamt = " & VTXTNetPPAmt & "," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'," & _
                       " remarks = " & VTXTRemarks & _
                       " where id = " & labid.Caption
        gconDMIS.Execute "update PMIS_PPDayTran set" & _
                       " trandate = " & VTXTPPDate & "," & _
                       " tranno = " & VTXTPPNo & _
                       " where tranno = '" & PrevPPNO & "'"

        LogAudit "E", "Purchase Processing", PrevPPNO
    End If
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "update PMIS_Counter set nextnumber = " & N2Str2Null(NewPPCunter) & ", lastupdate = '" & LOGDATE & "', usercode = " & N2Str2Null(LOGCODE) & " where modul = 'PP'"
    End If
    rsRefresh
    On Error Resume Next
    rsPP_HD.Find "id = " & labid.Caption
    cleargrid grdDetails
    FillDetails
    PP_TOTVAT = PP_TOTINVAMT - PP_TOTUCOST
    gconDMIS.Execute "update PMIS_PP_Hd set" & _
                   " PP_amount = " & PP_TOTUCOST & "," & _
                   " netPPamt = " & PP_TOTINVAMT & "," & _
                   " ds_desc1 = '" & "VAT" & "'," & _
                   " ds_amt1 = " & PP_TOTVAT & _
                   " where id = " & labid.Caption
    rsRefresh
    On Error Resume Next
    rsPP_HD.Find "id = " & labid.Caption
    cmdCancel.Value = True
    If AddorEdit = "ADD" Then
        cmdAddTran_Click
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If Picture1.Visible = True Then
                fraPrintFormat.Visible = False
                cmdPrintFormat.Visible = False
                cmdPrintFormat.ZOrder 0
                fraPrintFormat.ZOrder 0
                SendToBack
                StoreMemVars
            End If
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(rsPP_HD!Status) = "P" Then
                    MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
                ElseIf Null2String(rsPP_HD!Status) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change"
                Else
                    cmdAddTran_Click
                End If
            End If
        Case vbKeyF4
            If Null2String(rsPP_HD!Status) <> "P" And Null2String(rsPP_HD!Status) <> "C" Then
                grdDetails_DblClick
            End If
        Case vbKeyF5
            If Null2String(rsPP_HD!Status) <> "P" And Null2String(rsPP_HD!Status) <> "C" Then
                grdDetails_DblClick
                cmdTranDelete_Click
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    textSearch.Text = "":                             'Picture6.ZOrder 0
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    txtPartID.Text = ""
    InitMemvars
    rsRefresh
    If Not rsPP_HD.EOF And Not rsPP_HD.BOF Then rsPP_HD.MoveLast
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    Set rsPP_HD = New ADODB.Recordset
    rsPP_HD.Open "select * from PMIS_PP_Hd where ordertype = '" & ORDERTYPE & "' order by ppno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitMemvars()
    txtPPNo.Text = ""
    Set rsCunter = New ADODB.Recordset
    rsCunter.Open "select modul,nextnumber from PMIS_Counter where modul = 'PP'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCunter.EOF And Not rsCunter.BOF Then
        If ORDERTYPE = "S" Then
            txtPPNo.Text = Null2String(rsCunter!nextnumber)
        End If
    End If
    txtPPDate.Text = LOGDATE
    txtSupCode.Text = ""
    FillCboSupName
    txtSup_Addrs.Text = ""
    Filltxtshipto
    txtPP_Amount.Text = ""
    txtDS1.Text = ""
    txtDS_Desc1.Text = ""
    txtDS_Amt1.Text = 0
    txtPP_Amount.Text = 0
    txtNetPPAmt.Text = 0
    labPPsted.Visible = False
    txtRemarks.Text = "Pls Type Your Message Here!"
    cleargrid grdDetails
    InitGrid
    InitCbo
    InitParts
End Sub

Sub StoreMemVars()
    If Not rsPP_HD.EOF And Not rsPP_HD.BOF Then
        labid.Caption = rsPP_HD!ID
        txtPPNo.Text = Null2String(rsPP_HD!ppno)
        txtPPDate.Text = Null2String(rsPP_HD!ppdate)
        txtSupCode.Text = Null2String(rsPP_HD!SupCode)
        cboSupName.Text = Null2String(rsPP_HD!supname)
        txtSup_Addrs.Text = Null2String(rsPP_HD!sup_addrs)
        txtDealerCode.Text = Null2String(rsPP_HD!dealercode)
        Filltxtshipto2 (Null2String(rsPP_HD!dealercode))
        txtPP_Amount.Text = ToDoubleNumber(N2Str2Zero(rsPP_HD!pp_amount))
        txtDS1.Text = N2Str2IntZero(rsPP_HD!ds1)
        txtDS_Desc1.Text = Null2String(rsPP_HD!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(rsPP_HD!ds_amt1))
        txtNetPPAmt.Text = ToDoubleNumber(N2Str2Zero(rsPP_HD!netppamt))
        txtRemarks.Text = Null2String(rsPP_HD!remarks)
        If Null2String(rsPP_HD!Status) = "P" Then
            labPPsted.Visible = True
            labPPsted.Caption = "POSTED"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            If LOGLEVEL = "ADM" Then cmdUnPost.Enabled = True
            cmdPrint.Enabled = True
            If LOGLEVEL = "ADM" Then cmdCancelPP.Enabled = True
        ElseIf Null2String(rsPP_HD!Status) = "C" Then
            labPPsted.Visible = True
            labPPsted.Caption = "CANCELLED"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdUnPost.Enabled = False
            cmdPrint.Enabled = False
            cmdCancelPP.Enabled = False
        Else
            labPPsted.Visible = False
            labPPsted.Caption = ""
            cmdEdit.Enabled = True
            cmdPost.Enabled = True
            cmdUnPost.Enabled = False
            cmdPrint.Enabled = True
            If LOGLEVEL = "ADM" Then cmdCancelPP.Enabled = True
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
        .ColWidth(3) = 2200
        .ColWidth(4) = 500
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        .ColWidth(7) = 1000
        .ColWidth(8) = 1500
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
        .Text = "Inv. Amt."
        .Col = 6
        .Text = "Cost"
        .Col = 7
        .Text = "Total Amount"
        .Col = 8
        .Text = "Remarks"
    End With
End Sub

Sub FillDetails()
' On Error Resume Next
    Pcnt = 0
    PP_TOTUCOST = 0
    PP_TOTINVAMT = 0

    Set rsPPDaytran = New ADODB.Recordset
    rsPPDaytran.Open "select id,itemno,PART_ORD,PART_SUP,tranqty,traninvamt,tranucost,tremarks from PMIS_PPDayTran where tranno = " & N2Str2Null(rsPP_HD!ppno) & " AND trantype = 'PP' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPPDaytran.EOF And Not rsPPDaytran.BOF Then
        Screen.MousePointer = 11
        rsPPDaytran.MoveFirst
        Do While Not rsPPDaytran.EOF
            Pcnt = Pcnt + 1
            grdDetails.AddItem rsPPDaytran!ID & Chr(9) & Null2String(rsPPDaytran!itemno) & Chr(9) & _
                               Null2String(rsPPDaytran!PART_ORD) & Chr(9) & _
                               Null2String(rsPPDaytran!PART_SUP) & Chr(9) & _
                               N2Str2IntZero(rsPPDaytran!tranqty) & Chr(9) & _
                               N2Str2Zero(rsPPDaytran!TRANINVAMT) & Chr(9) & _
                               N2Str2Zero(rsPPDaytran!TRANUCOST) & Chr(9) & _
                               Format(N2Str2Zero(rsPPDaytran!tranqty) * N2Str2Zero(rsPPDaytran!TRANUCOST), MAXIMUM_DIGIT) & Chr(9) & _
                               Null2String(rsPPDaytran!TREMARKS)
            PP_TOTUCOST = PP_TOTUCOST + (N2Str2Zero(rsPPDaytran!tranqty) * N2Str2Zero(rsPPDaytran!TRANUCOST))
            PP_TOTINVAMT = PP_TOTINVAMT + (N2Str2Zero(rsPPDaytran!tranqty) * N2Str2Zero(rsPPDaytran!TRANINVAMT))
            rsPPDaytran.MoveNext
        Loop
        If Pcnt <> 0 Then grdDetails.RemoveItem 1
        PP_TOTVAT = (PP_TOTUCOST * ConvertToBIRDecimalFormat(VAT_RATE)) - PP_TOTUCOST
        If NumericVal(PP_TOTVAT) <> 0 Then
            txtDS1.Text = VAT_RATE
            txtDS_Desc1.Text = "VAT"
            txtDS_Amt1.Text = ToDoubleNumber(PP_TOTVAT)
            txtNetPPAmt.Text = ToDoubleNumber(NumericVal(txtPP_Amount.Text) + NumericVal(txtDS_Amt1.Text))
        End If
        txtPP_Amount.Text = Format(txtPP_Amount.Text, MAXIMUM_DIGIT)
        Screen.MousePointer = 0
    Else
        cleargrid grdDetails
    End If
End Sub

Function SetSTOCKDESC(ppp As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select STOCKNO,STOCKDESC from PMIS_STOCKMAS where STOCKNO= '" & ppp & "'", gconDMIS
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetSTOCKDESC = Null2String(rsPartMas!STOCKDESC)
    End If
End Function

Function SetSTOCKDESC2(pid As Variant)
    Dim rsDNPP                         As ADODB.Recordset
    Set rsDNPP = New ADODB.Recordset
    rsDNPP.Open "Select id,descriptio from PMIS_DNPP where id = " & N2Str2IntZero(pid), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsDNPP.EOF And Not rsDNPP.BOF Then
        SetSTOCKDESC2 = Null2String(rsDNPP!DESCRIPTIO)
    End If
End Function

Function SetSTOCKNO(pid As Variant)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select id,STOCKNO from PMIS_STOCKMAS where id = " & N2Str2IntZero(pid), gconDMIS
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetSTOCKNO = Null2String(rsPartMas!STOCKNO)
    End If
End Function

Function SetPartIDSTOCKNO(DDD As String)
    Dim rsDNPP                         As ADODB.Recordset
    Set rsDNPP = New ADODB.Recordset
    rsDNPP.Open "Select id,STOCKNUMBER from PMIS_DNPP where STOCKNUMBER = " & N2Str2Null(DDD), gconDMIS
    If Not rsDNPP.EOF And Not rsDNPP.BOF Then
        SetPartIDSTOCKNO = Null2String(rsDNPP!ID)
    End If
End Function

Function SetPartIDDesc(DDD As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select id,STOCKDESC from PMIS_STOCKMAS where ltrim(rtrim(STOCKDESC))) = " & N2Str2Null(DDD), gconDMIS
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartIDDesc = Null2String(rsPartMas!ID)
    End If
End Function

Function SetPartPrice(ppp As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select STOCKNO,mac from PMIS_STOCKMAS where STOCKNO = " & N2Str2Null(ppp), gconDMIS
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartPrice = Null2String(rsPartMas!Mac)
    End If
End Function

Sub FillCboSupName()
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "select supname from PMIS_vw_Supplier ORDER BY SUPNAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        rsSupplier.MoveFirst
        cboSupName.Clear
        Do While Not rsSupplier.EOF
            cboSupName.AddItem Null2String(rsSupplier!supname)
            rsSupplier.MoveNext
        Loop
    End If
End Sub

Sub Filltxtshipto()
    Set rsProfile = New ADODB.Recordset
    rsProfile.Open "select * from ALL_Profile", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        txtDealerCode.Text = Null2String(rsProfile!companycode)
        txtShipTo.Text = Null2String(rsProfile!CompanyName)
        txtSHP_addrs.Text = Null2String(rsProfile!Companyaddress)
    End If
End Sub

Sub Filltxtshipto2(param As String)
    Set rsProfile = New ADODB.Recordset
    rsProfile.Open "select * from ALL_Profile where companycode = '" & param & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        txtDealerCode.Text = Null2String(rsProfile!companycode)
        txtShipTo.Text = Null2String(rsProfile!CompanyName)
        txtSHP_addrs.Text = Null2String(rsProfile!Companyaddress)
    End If
End Sub

Function SetSupdesc(ppp As String)
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "Select supcode,supname,sup_addrs from PMIS_vw_Supplier where supcode = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        SetSupdesc = Null2String(rsSupplier!supname)
        txtSup_Addrs.Text = Null2String(rsSupplier!sup_addrs)
    Else
        txtSup_Addrs.Text = ""
    End If
End Function

Function SetSupCode(nnn As String)
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "Select supname,supcode,sup_addrs from PMIS_vw_Supplier where supname = '" & nnn & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        SetSupCode = Null2String(rsSupplier!SupCode)
        txtSup_Addrs.Text = Null2String(rsSupplier!sup_addrs)
    Else
        txtSup_Addrs.Text = ""
    End If
End Function

Sub InitParts()
    txtTranItemNo.Text = Format(Pcnt + 1, "0000")
    cboTranPartNo.Text = ""
    cboTranDescription.Text = ""
    txtUnit.Text = ""
    txtTRemarks.Text = ""
    txtTranQty.Text = 1
    txtTranINVAmt.Text = "0.00"
    txtTranTotalAmt.Text = "0.00"
End Sub

Function StorePartsEntry(ByVal ID As Variant)
    Set rsPPDaytran = New ADODB.Recordset
    rsPPDaytran.Open "select id,itemno,PART_ORD,PART_SUP,tranqty,traninvamt,tranucost,unit,tremarks from PMIS_PPDayTran where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPPDaytran.EOF And Not rsPPDaytran.BOF Then
        labDetID.Caption = rsPPDaytran!ID
        txtTranItemNo.Text = Null2String(rsPPDaytran!itemno)
        cboTranPartNo.Text = Null2String(rsPPDaytran!PART_ORD)
        cboTranDescription.Text = Null2String(rsPPDaytran!PART_SUP)
        txtTranQty.Text = N2Str2IntZero(rsPPDaytran!tranqty)
        txtTranINVAmt.Text = ToDoubleNumber(N2Str2Zero(rsPPDaytran!TRANINVAMT))
        txtUnitCost.Text = ToDoubleNumber(N2Str2Zero(rsPPDaytran!TRANUCOST))
        txtTranTotalAmt.Text = ToDoubleNumber(N2Str2IntZero(rsPPDaytran!tranqty) * N2Str2Zero(rsPPDaytran!TRANINVAMT))
        txtUnit.Text = Null2String(rsPPDaytran!unit)
        txtTRemarks.Text = Null2String(rsPPDaytran!TREMARKS)
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISPurchProcessing = Nothing
    UnloadForm Me
End Sub

Private Sub grdDetails_DblClick()
    If Null2String(rsPP_HD!Status) = "P" Then
        MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
    ElseIf Null2String(rsPP_HD!Status) = "C" Then
        MsgSpeechBox "Transactions are Already Cancelled and cannot be Change"
    Else
        Dim Fild                       As String
        grdDetails.Row = grdDetails.Row
        grdDetails.Col = 0
        Fild = grdDetails.Text
        If Fild <> "" And Fild <> "No Entry" Then
            AddorEdit = "EDIT"
            BringToFront
            fraAddTran.Caption = "Edit Parts"
            StorePartsEntry (Fild)
        Else
            MsgSpeechBox "No Entry on Parts"
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
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "select STOCKNO,STOCKDESC from PMIS_STOCKMAS  order BY STOCKDESC ASC", gconDMIS
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        rsPartMas.MoveFirst
        cboTranPartNo.Clear
        cboTranDescription.Clear
        Do While Not rsPartMas.EOF
            cboTranPartNo.AddItem Null2String(rsPartMas!STOCKNO)
            cboTranDescription.AddItem Null2String(rsPartMas!STOCKDESC)
            rsPartMas.MoveNext
        Loop
    End If
End Sub

Private Sub cboTranDescription_LostFocus()
    cboTranDescription.Text = cboTranDescription.Text
End Sub

Private Sub grdDetails_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Timer1_Timer()
    If labPPsted.Caption <> "" Then
        If labPPsted.Visible = True Then
            labPPsted.Visible = False
        Else
            labPPsted.Visible = True
        End If
    End If
End Sub

Private Sub txtPPDate_LostFocus()
    txtPPDate.Text = Format(txtPPDate.Text, "SHORT DATE")
End Sub

Private Sub txtPPNo_LostFocus()
    If Len(txtPPNo.Text) >= 3 Then
        If AddorEdit = "ADD" Then
            Dim rsPP_HDDup             As ADODB.Recordset
            Set rsPP_HDDup = New ADODB.Recordset
            rsPP_HDDup.Open "select ppno from PMIS_PP_Hd where PPno = '" & txtPPNo.Text & "'", gconDMIS
            If Not rsPP_HDDup.EOF And Not rsPP_HDDup.BOF Then
                MsgSpeechBox "PP Number Already Exist!"
            End If
        End If
    End If
End Sub

Private Sub txtRemarks_GotFocus()
    MsgSpeech "Pls Type Your Message Here!"
    If txtRemarks.Text = "Pls Type Your Message Here!" Then txtRemarks.Text = ""
End Sub

Private Sub txtSupCode_Change()
    cboSupName.Text = SetSupdesc(txtSupCode.Text)
End Sub

Private Sub txttranQty_Change()
    If txtTranQty.Text <> "" Then
        txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE))
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtUnitCost.Text))
    End If
End Sub

Private Sub txtTranINVAmt_Change()
    If txtTranINVAmt.Text <> "" Then
        txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE))
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtUnitCost.Text))
    End If
End Sub

Private Sub txtTranINVAmt_GotFocus()
    If NumericVal(txtTranINVAmt.Text) = 0 Then txtTranINVAmt.Text = ""
End Sub

Private Sub txtTranINVAmt_LostFocus()
    If txtTranINVAmt.Text = "" Then txtTranINVAmt.Text = 0
    txtTranINVAmt.Text = ToDoubleNumber(txtTranINVAmt.Text)
End Sub

Private Sub txtTranQty_GotFocus()
    If NumericVal(txtTranQty.Text) = 1 Then txtTranQty.Text = ""
End Sub

Private Sub txtTranQty_LostFocus()
    If NumericVal(txtTranQty.Text) = 0 Then txtTranQty.Text = 1
    txtTranQty.Text = Format(txtTranQty.Text, DIGIT_FORMAT)
End Sub

Private Sub txtTranTotalAmt_Change()
    txtTranTotalAmt.Text = Format(txtTranTotalAmt.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtUnitCost_LostFocus()
    txtUnitCost.Text = Format(txtUnitCost.Text, MAXIMUM_DIGIT)
End Sub

'SEARCH MODULE
Private Sub lstPP_HD_GotFocus()
    rsPP_HD.Bookmark = rsFind(rsPP_HD.Clone, "ID", lstPP_HD.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstPP_HD_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If optPPNo.Value = True Then
        rsPP_HD.Bookmark = rsFind(rsPP_HD.Clone, "ppno", Item).Bookmark
    Else
        rsPP_HD.Bookmark = rsFind(rsPP_HD.Clone, "ID", lstPP_HD.SelectedItem.SubItems(1)).Bookmark
    End If
    StoreMemVars
End Sub

Private Sub lstPP_HD_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPP_HD
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

Private Sub lstPP_HD_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstPP_HD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If optPPNo.Value = True Then
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
    Dim rsPP_HD                        As ADODB.Recordset
    lstPP_HD.Sorted = False: lstPP_HD.ListItems.Clear
    Set rsPP_HD = New ADODB.Recordset
    Set rsPP_HD = gconDMIS.Execute("select ppno,ID from PMIS_PP_Hd where ordertype = '" & ORDERTYPE & "' order by ppno asc")
    If Not (rsPP_HD.EOF And rsPP_HD.BOF) Then
        lstPP_HD.Enabled = True
        Listview_Loadval Me.lstPP_HD.ListItems, rsPP_HD
        lstPP_HD.Refresh
    Else
        lstPP_HD.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsPP_HD                        As ADODB.Recordset
    lstPP_HD.Sorted = False: lstPP_HD.ListItems.Clear
    Set rsPP_HD = New ADODB.Recordset
    Set rsPP_HD = gconDMIS.Execute("select ppno, ID from PMIS_PP_Hd where ordertype = '" & ORDERTYPE & "' and ppno like'" & XXX & "%'")
    If Not (rsPP_HD.EOF And rsPP_HD.BOF) Then
        lstPP_HD.Enabled = True
        Listview_Loadval Me.lstPP_HD.ListItems, rsPP_HD
        lstPP_HD.Refresh
    Else
        lstPP_HD.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim rsPP_HD                        As ADODB.Recordset
    lstPP_HD.Sorted = False: lstPP_HD.ListItems.Clear
    Set rsPP_HD = New ADODB.Recordset
    Set rsPP_HD = gconDMIS.Execute("select supname, ID from PMIS_PP_Hd where ordertype = '" & ORDERTYPE & "' order by ppno asc")
    If Not (rsPP_HD.EOF And rsPP_HD.BOF) Then
        lstPP_HD.Enabled = True
        Listview_Loadval Me.lstPP_HD.ListItems, rsPP_HD
        lstPP_HD.Refresh
    Else
        lstPP_HD.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsPP_HD                        As ADODB.Recordset
    lstPP_HD.Sorted = False: lstPP_HD.ListItems.Clear
    Set rsPP_HD = New ADODB.Recordset
    Set rsPP_HD = gconDMIS.Execute("select supname, ID from PMIS_PP_Hd where ordertype = '" & ORDERTYPE & "' and supname like '" & XXX & "%' order by ppno asc")
    If Not (rsPP_HD.EOF And rsPP_HD.BOF) Then
        lstPP_HD.Enabled = True
        Listview_Loadval Me.lstPP_HD.ListItems, rsPP_HD
        lstPP_HD.Refresh
    Else
        lstPP_HD.Enabled = False
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then lstPP_HD.SetFocus
End Sub

Private Sub optRONo_Click()
    lstPP_HD.ColumnHeaders(1).Text = "Sup. Name"
    lstPP_HD.ColumnHeaders(1).Width = 4000
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    textSearch.SetFocus
End Sub

Private Sub optPPNo_Click()
    lstPP_HD.ColumnHeaders(1).Text = "Tran. No."
    lstPP_HD.ColumnHeaders(1).Width = 2150
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    textSearch.SetFocus
End Sub
