VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Begin VB.Form frmSMIS_Trans_Ordering 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Purchase Order"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Ordering.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   10950
   Begin VB.PictureBox picContact 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   7740
      ScaleHeight     =   1275
      ScaleWidth      =   2985
      TabIndex        =   88
      Top             =   4980
      Width           =   3015
      Begin VB.CommandButton cmdpicCancel 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   1560
         TabIndex        =   94
         Top             =   750
         Width           =   1215
      End
      Begin VB.TextBox txtcontact 
         Height          =   315
         Left            =   60
         TabIndex        =   90
         Top             =   360
         Width           =   2865
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "Ok"
         Height          =   405
         Left            =   360
         TabIndex        =   89
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
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
         Left            =   240
         TabIndex        =   93
         Top             =   30
         Width           =   2445
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   285
         Left            =   0
         TabIndex        =   92
         Top             =   0
         Width           =   3105
         _Version        =   655364
         _ExtentX        =   5477
         _ExtentY        =   503
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   30
         TabIndex        =   91
         Top             =   30
         Width           =   2295
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   0
      ScaleHeight     =   7335
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.Frame fraSearch 
         BorderStyle     =   0  'None
         Height          =   7365
         Left            =   30
         TabIndex        =   1
         Top             =   0
         Width           =   2925
         Begin VB.TextBox txtPODetailID 
            Alignment       =   2  'Center
            Height          =   495
            Left            =   1725
            TabIndex        =   6
            Text            =   "0"
            Top             =   180
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox txtID 
            Height          =   495
            Left            =   1710
            TabIndex        =   3
            Top             =   210
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.OptionButton optDate 
            Caption         =   "D&ate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1710
            TabIndex        =   7
            Top             =   300
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.OptionButton optPO 
            Caption         =   "&PO Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   4
            Top             =   390
            Value           =   -1  'True
            Width           =   1845
         End
         Begin VB.OptionButton optVModel 
            Caption         =   "&Vehicle Model[Description]"
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
            Left            =   90
            TabIndex        =   5
            Top             =   660
            Width           =   2745
         End
         Begin VB.TextBox textSearch 
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
            Height          =   360
            Left            =   90
            MaxLength       =   35
            TabIndex        =   8
            Text            =   "TEXT"
            Top             =   990
            Width           =   2715
         End
         Begin MSComctlLib.ListView lstPO 
            Height          =   5895
            Left            =   90
            TabIndex        =   9
            Top             =   1380
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   10398
            View            =   3
            LabelEdit       =   1
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
            MouseIcon       =   "Ordering.frx":08CA
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Date"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PO"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   2540
            EndProperty
         End
         Begin Crystal.CrystalReport rptPO 
            Left            =   1350
            Top             =   1260
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.Label Label22 
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
            TabIndex        =   2
            Top             =   150
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox picTop 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   3030
      ScaleHeight     =   6375
      ScaleWidth      =   7875
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   7875
      Begin VB.CommandButton cmdAddFromPO 
         Height          =   345
         Left            =   1860
         Picture         =   "Ordering.frx":0A2C
         Style           =   1  'Graphical
         TabIndex        =   86
         TabStop         =   0   'False
         ToolTipText     =   "Make Purchase Order (VPO) Reference to Customer Sales Order"
         Top             =   390
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton Command6 
         Caption         =   "::"
         Height          =   375
         Left            =   7470
         TabIndex        =   85
         ToolTipText     =   "Edit Transaction Date"
         Top             =   540
         Width           =   345
      End
      Begin VB.CommandButton Command5 
         Caption         =   "::"
         Height          =   375
         Left            =   7470
         TabIndex        =   84
         ToolTipText     =   "Edit Transaction Date"
         Top             =   120
         Width           =   345
      End
      Begin VB.TextBox txtModelCode 
         Height          =   345
         Left            =   4200
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1290
         Width           =   1485
      End
      Begin VB.ComboBox cboSource 
         Height          =   315
         ItemData        =   "Ordering.frx":0BF6
         Left            =   2970
         List            =   "Ordering.frx":0BF8
         TabIndex        =   35
         Text            =   "Combo1"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtDatePO 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5700
         TabIndex        =   14
         Top             =   120
         Width           =   1755
      End
      Begin VB.TextBox txtDueDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5700
         TabIndex        =   17
         Top             =   540
         Width           =   1755
      End
      Begin VB.TextBox txtPONO 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   150
         MaxLength       =   7
         TabIndex        =   15
         Top             =   360
         Width           =   1725
      End
      Begin VB.Frame fraCheckDetail 
         BorderStyle     =   0  'None
         Height          =   2685
         Left            =   3420
         TabIndex        =   46
         Top             =   2280
         Width           =   4305
         Begin VB.TextBox TXTWTAX 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   1710
            MaxLength       =   11
            TabIndex        =   57
            Text            =   "0.00"
            Top             =   2250
            Width           =   2505
         End
         Begin VB.CheckBox Check1 
            Caption         =   "W/H Tax"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   720
            TabIndex        =   58
            Top             =   2280
            Width           =   1845
         End
         Begin VB.TextBox txtSubsidy 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   1710
            TabIndex        =   56
            Top             =   1800
            Width           =   2505
         End
         Begin VB.TextBox txtPy_CD_Amount 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   1710
            TabIndex        =   54
            Top             =   1350
            Width           =   2505
         End
         Begin VB.TextBox txtPy_CD_CheckNo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1710
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   540
            Width           =   2505
         End
         Begin VB.ComboBox cboPy_CD_BankName 
            Appearance      =   0  'Flat
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
            ItemData        =   "Ordering.frx":0BFA
            Left            =   1710
            List            =   "Ordering.frx":0BFC
            TabIndex        =   48
            Text            =   "Combo1"
            Top             =   120
            Width           =   2505
         End
         Begin MSComCtl2.DTPicker txPy_CD_Date 
            Height          =   345
            Left            =   1710
            TabIndex        =   52
            Top             =   930
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   133693441
            CurrentDate     =   39248
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Subsidy"
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
            Height          =   225
            Left            =   990
            TabIndex        =   55
            Top             =   1920
            Width           =   675
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Height          =   225
            Left            =   1275
            TabIndex        =   51
            Top             =   960
            Width           =   390
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check No."
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
            Height          =   225
            Left            =   810
            TabIndex        =   49
            Top             =   510
            Width           =   855
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
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
            Height          =   225
            Left            =   1200
            TabIndex        =   47
            Top             =   150
            Width           =   435
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
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
            Height          =   225
            Left            =   1005
            TabIndex        =   53
            Top             =   1410
            Width           =   660
         End
      End
      Begin VB.Frame fraCrNo 
         BorderStyle     =   0  'None
         Caption         =   "Financing Option"
         Height          =   1725
         Left            =   -30
         TabIndex        =   40
         Top             =   2940
         Width           =   3345
         Begin VB.ComboBox cboPy_FinLcIssuingBank 
            Appearance      =   0  'Flat
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
            ItemData        =   "Ordering.frx":0BFE
            Left            =   120
            List            =   "Ordering.frx":0C00
            TabIndex        =   42
            Text            =   "Combo1"
            Top             =   255
            Width           =   2970
         End
         Begin VB.TextBox txtPy_LCNo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   900
            Width           =   2985
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Financing /LC Issuing Bank"
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
            Height          =   225
            Left            =   120
            TabIndex        =   41
            Top             =   0
            Width           =   2265
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Letter of Credit Number"
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
            Height          =   225
            Left            =   120
            TabIndex        =   43
            Top             =   645
            Width           =   1995
         End
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   7800
         ScaleHeight     =   195
         ScaleWidth      =   7605
         TabIndex        =   18
         Top             =   990
         Width           =   7605
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F3 - Add "
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   1530
            TabIndex        =   19
            Top             =   0
            Width           =   675
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F4 - Edit "
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   2310
            TabIndex        =   20
            Top             =   0
            Width           =   630
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "F12 - Un-Post Transaction"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   5700
            TabIndex        =   23
            Top             =   0
            Width           =   2445
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "F8 - Post Transaction"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   4050
            TabIndex        =   22
            Top             =   0
            Width           =   1905
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F5 - Delete"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   3090
            TabIndex        =   21
            Top             =   30
            Width           =   780
         End
      End
      Begin VB.ComboBox cboModeOfPayment 
         Appearance      =   0  'Flat
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
         ItemData        =   "Ordering.frx":0C02
         Left            =   90
         List            =   "Ordering.frx":0C15
         TabIndex        =   39
         Text            =   "Combo1"
         Top             =   2550
         Width           =   3000
      End
      Begin VB.CommandButton Command1 
         Height          =   345
         Left            =   4530
         Picture         =   "Ordering.frx":0C5F
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Select Customer"
         Top             =   6390
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtCusCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   6390
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtFuel 
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
         Left            =   6240
         TabIndex        =   37
         Top             =   1920
         Width           =   1365
      End
      Begin VB.TextBox txtModelYear 
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
         Left            =   4800
         TabIndex        =   36
         Top             =   1920
         Width           =   1395
      End
      Begin VB.TextBox txtModel 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1290
         Width           =   1965
      End
      Begin VB.ComboBox cboModelDescript 
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
         Height          =   345
         Left            =   90
         TabIndex        =   27
         Text            =   "txtDescript"
         Top             =   1290
         Width           =   4065
      End
      Begin VB.ComboBox cboColor 
         Appearance      =   0  'Flat
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
         Left            =   90
         TabIndex        =   34
         Text            =   "Combo1"
         Top             =   1920
         Width           =   2865
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   30
         Top             =   6570
      End
      Begin VB.TextBox txtNotes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   90
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   59
         Top             =   4980
         Width           =   7635
      End
      Begin VB.Label LABALLOWREPRINT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   2310
         TabIndex        =   87
         Top             =   660
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5190
         TabIndex        =   13
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Due Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4860
         TabIndex        =   16
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "POSTED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1380
         TabIndex        =   12
         Top             =   0
         Width           =   3825
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "PO No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   11
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mode of Payment "
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
         Height          =   225
         Left            =   90
         TabIndex        =   38
         Top             =   2310
         Width           =   1515
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cus Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1995
         TabIndex        =   60
         Top             =   6450
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2970
         TabIndex        =   31
         Top             =   1695
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fuel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6240
         TabIndex        =   33
         Top             =   1695
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4800
         TabIndex        =   32
         Top             =   1695
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
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
         Height          =   225
         Left            =   4200
         TabIndex        =   25
         Top             =   1020
         Width           =   990
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Description"
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
         Height          =   225
         Left            =   90
         TabIndex        =   24
         Top             =   1020
         Width           =   1530
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Height          =   225
         Left            =   5730
         TabIndex        =   26
         Top             =   1020
         Width           =   510
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
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
         Height          =   225
         Left            =   90
         TabIndex        =   30
         Top             =   1695
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   45
         Top             =   4770
         Width           =   495
      End
   End
   Begin VB.PictureBox picmultiple 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   5460
      ScaleHeight     =   2055
      ScaleWidth      =   2445
      TabIndex        =   63
      Top             =   2040
      Visible         =   0   'False
      Width           =   2475
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   780
         ScaleHeight     =   885
         ScaleWidth      =   2190
         TabIndex        =   66
         Top             =   1080
         Width           =   2190
         Begin VB.CommandButton Command3 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   840
            MouseIcon       =   "Ordering.frx":0E29
            MousePointer    =   99  'Custom
            Picture         =   "Ordering.frx":0F7B
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Save"
            Height          =   795
            Left            =   150
            MouseIcon       =   "Ordering.frx":12B9
            MousePointer    =   99  'Custom
            Picture         =   "Ordering.frx":140B
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   60
            Width           =   705
         End
      End
      Begin VB.TextBox txtMultiplePONo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   360
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   480
         Width           =   1725
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   -30
         TabIndex        =   64
         Top             =   0
         Width           =   2565
         _Version        =   655364
         _ExtentX        =   4524
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   " Enter Number Quantity"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   2940
      ScaleHeight     =   915
      ScaleWidth      =   8490
      TabIndex        =   72
      Top             =   6390
      Width           =   8490
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   7140
         MouseIcon       =   "Ordering.frx":175B
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":18AD
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   6450
         MouseIcon       =   "Ordering.frx":1C13
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":1D65
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdCancelCO 
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
         Left            =   5760
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Ordering.frx":20CB
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":221D
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Cancel this Transaction"
         Top             =   60
         Width           =   705
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
         Left            =   5070
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Ordering.frx":2557
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":26A9
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Post this Transaction"
         Top             =   60
         Width           =   705
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
         Left            =   4380
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Ordering.frx":29CE
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":2B20
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Unpost this Transaction"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   3690
         MouseIcon       =   "Ordering.frx":2E65
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":2FB7
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   3000
         MouseIcon       =   "Ordering.frx":3313
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":3465
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add &Multiple"
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
         Left            =   2310
         MouseIcon       =   "Ordering.frx":3778
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":38CA
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   1620
         MouseIcon       =   "Ordering.frx":3BDD
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":3D2F
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   930
         MouseIcon       =   "Ordering.frx":4029
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":417B
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   240
         MouseIcon       =   "Ordering.frx":44D3
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":4625
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   9300
      ScaleHeight     =   885
      ScaleWidth      =   2190
      TabIndex        =   69
      Top             =   6390
      Width           =   2190
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   780
         MouseIcon       =   "Ordering.frx":4984
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":4AD6
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   90
         MouseIcon       =   "Ordering.frx":4E14
         MousePointer    =   99  'Custom
         Picture         =   "Ordering.frx":4F66
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   60
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmSMIS_Trans_Ordering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPO                                                              As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim DontChange                                                        As Boolean
Dim rsParts                                                           As ADODB.Recordset
Dim rsS_Model                                                         As ADODB.Recordset
Dim rsColor                                                           As ADODB.Recordset
Dim WithEvents SearchMaster                                           As frmSMIS_Mis_SearchMaster
Attribute SearchMaster.VB_VarHelpID = -1
Dim MULTIPLEPO                                                        As Boolean

Function GetModeOfPayment(XXX)
    If XXX = "LC" Then
        GetModeOfPayment = "Letter of Credit"
    ElseIf XXX = "CA" Then
        GetModeOfPayment = "Cash"
    ElseIf XXX = "OA" Then
        GetModeOfPayment = "Open Account"
    ElseIf XXX = "PN" Then
        GetModeOfPayment = "Promissory Note"
    ElseIf XXX = "FC" Then
        GetModeOfPayment = "Financing Co."
    End If
End Function

Function SetModeOfPayment(XXX)
    XXX = UCase(XXX)
    If XXX = UCase("Letter of Credit") Then
        SetModeOfPayment = "LC"
    ElseIf XXX = UCase("Open Account") Then
        SetModeOfPayment = "OA"
    ElseIf XXX = UCase("Promissory Note") Then
        SetModeOfPayment = "PN"
    ElseIf XXX = UCase("Financing Co.") Then
        SetModeOfPayment = "FC"
    ElseIf XXX = UCase("Cash") Then
        SetModeOfPayment = "CA"
    End If
End Function

Function SetStatus(XString) As String
    'when 'FO' then 'FOR ORDERING'
    'when 'BO' then 'BACK ORDER STAGE'
    'when 'AS' then 'ALLOCATION STAGE'
    'when 'KS' then 'PICKING STAGE'
    'when 'PS' then 'PACKING STAGE'
    'when 'SS' then 'SHIPPING STAGE'

    Select Case XString
        Case "FO"
            SetStatus = "FOR ORDERING"
        Case "BO"
            SetStatus = "BACK ORDER STAGE"
        Case "AS"
            SetStatus = "ALLOCATION STAGE"
        Case "KS"
            SetStatus = "PICKING STAGE"
        Case "PS"
            SetStatus = "PACKING STAGE"
        Case "SS"
            SetStatus = "SHIPPING STAGE"
    End Select
End Function

Sub CboRefresh()
    Dim rsTemp                                                        As ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("Select DISTINCT(FinLcIssuingBank) from SMIS_PO WHERE LEN(FinLcIssuingBank)>0 Order by 1 asc")

    Combo_Loadval cboPy_FinLcIssuingBank, rsTemp

    Set rsTemp = gconDMIS.Execute("Select DISTINCT(CD_BankName) from SMIS_PO  WHERE LEN(CD_BankName)>0 Order by 1 asc")

    Combo_Loadval cboPy_CD_BankName, rsTemp

End Sub

Sub FillSearchGrid()
    Dim temprs                                                        As ADODB.Recordset
    lstPO.Enabled = False
    Dim XXX                                                           As String
    If optVModel.Value = True Then
        Set temprs = gconDMIS.Execute("SELECT ModelDescript, PO_NO , ID FROM SMIS_PO WHERE ModelDescript Like " & N2Str2Null(ReplaceQuote(textSearch & "%")) & " ORDER BY PO_NO DESC")
    ElseIf optDate.Value = True Then
        Set temprs = gconDMIS.Execute("SELECT convert(varchar, DateOrdered,101), PO_NO , ID FROM SMIS_PO WHERE  convert(varchar, DateOrdered,101)  Like " & N2Str2Null(ReplaceQuote(textSearch & "%")) & " ORDER BY PO_NO DESC")
    ElseIf optPO.Value = True Then
        XXX = Format(textSearch, "000000")
        Set temprs = gconDMIS.Execute("SELECT convert(varchar, DateOrdered,101), PO_NO ,  ID FROM SMIS_PO WHERE  PO_NO Like '%" & XXX & "%' ORDER BY PO_NO DESC")
    End If


    If Not (temprs.EOF Or temprs.BOF) Then
        flex_FillListView temprs, lstPO

        'Listview_Loadval lstPO.ListItems, Temprs
        lstPO.Enabled = True
    End If

    Set temprs = Nothing
End Sub

Sub InitCombo()
    Dim SQL                                                           As String

    SQL = "Select AccessoriesName , ID from SMIS_VACC "

    Set rsParts = New ADODB.Recordset
    Call rsParts.Open(SQL, gconDMIS, adOpenKeyset, adLockReadOnly)

    Set rsS_Model = New ADODB.Recordset
    Call rsS_Model.Open("Select descript from All_Model where LEN(code)<> 0 order by descript asc", gconDMIS, adOpenKeyset)
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboModelDescript.Clear
        Do While Not rsS_Model.EOF
            cboModelDescript.AddItem UCase(Null2String(rsS_Model!DESCRIPT))
            rsS_Model.MoveNext
        Loop
    End If
    Set rsColor = New ADODB.Recordset
    Call rsColor.Open("Select DISTINCT(Color_Desc) as Color_Desc from All_Color  order by 1 asc", gconDMIS, adOpenKeyset)
    If Not rsColor.EOF And Not rsColor.BOF Then
        rsColor.MoveFirst
        cboColor.Clear
        Do While Not rsColor.EOF
            cboColor.AddItem UCase(Null2String(rsColor!color_desc))
            rsColor.MoveNext
        Loop
    End If
    Dim rsSource                                                      As ADODB.Recordset
    Set rsSource = New ADODB.Recordset
    Call rsSource.Open("SELECT DISTINCT(SOURCE)  as source FROM SMIS_PO WHERE SOURCE IS NOT NULL  ORDER BY 1 ASC", gconDMIS, adOpenKeyset)
    If Not rsSource.EOF And Not rsSource.BOF Then
        rsSource.MoveFirst
        cboSource.Clear
        Do While Not rsColor.EOF
            cboSource.AddItem UCase(Null2String(rsSource!Source))
            rsSource.MoveNext
        Loop
    End If

End Sub

Sub initMemvars()
    Dim cntrl                                                         As Control
    LABALLOWREPRINT = ""
    For Each cntrl In Me.ControlS
        If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
            cntrl.Text = vbNullString
        End If
    Next
    cboModelDescript = ""
    txtPODetailID = 0
    txtID = 0
    lblSTATUS = ""
    txtMultiplePONo = 1
    txtPy_CD_Amount = "0.00"
    cboSource = "Mitsubishi"
    txtDatePO.Enabled = True: txtDueDate.Enabled = True
End Sub

Sub rsRefresh()
    Set rsPO = New ADODB.Recordset
    Call rsPO.Open("SELECT  * FROM SMIS_PO order by PO_NO ASC ", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub

Sub SearchID(XXX)
    If Not (rsPO.EOF Or rsPO.BOF) Then
        rsPO.Find ("ID=" & XXX)
        StoreMemVars

    End If

End Sub

Sub SetModelLine(XXX As String, ByCode As Boolean)
    Dim temprs                                                        As ADODB.Recordset
    txtModelCode = ""
    txtPy_CD_Amount = ""
    cmdSave.Enabled = False
    If ByCode = True Then
        Set temprs = gconDMIS.Execute("Select DESCRIPT, MODEL, CODE ,COSTPRICE from ALL_MODEL WHERE CODE=" & N2Str2Null(ReplaceQuote(XXX)))
    Else
        Set temprs = gconDMIS.Execute("Select DESCRIPT, MODEL , CODE ,COSTPRICE from ALL_MODEL WHERE DESCRIPT=" & N2Str2Null(ReplaceQuote(XXX)))
    End If
    DontChange = True

    If Not (temprs.EOF Or temprs.BOF) Then
        If ByCode = True Then
            cboModelDescript.Text = UCase(Null2String(temprs!DESCRIPT))
        Else
            txtModelCode = UCase(Null2String(temprs!CODE))
        End If

        txtPy_CD_Amount = FormatNumber(NumericVal(temprs!costprice), 2)
        txtModel = UCase(Null2String(temprs!Model))
        cmdSave.Enabled = True
    Else
        MsgBox " Invalid Model Code ! " & vbCrLf & "Try Again Or Select From Drop Down of Model Description", vbInformation
        If ByCode = True Then
            On Error Resume Next
            txtModelCode.SetFocus
        Else
            On Error Resume Next
            cboModelDescript.SetFocus
        End If
        cmdSave.Enabled = False
    End If
End Sub

Sub StoreMemVars()
    Dim MRRSTATUS                                                     As ADODB.Recordset
    If Not (rsPO.EOF Or rsPO.BOF) Then
        cmdEdit.Enabled = True
        LABALLOWREPRINT = Null2String(rsPO!PRINTED)
        'ID, PO_NO, DateOrdered, Descript, Model, YearModel, Source, Color, Fuel, DatePullOut, DateReleased, DateInvoiced, CustomerCode, Status, Notes FROM
        cboSource = Null2String(rsPO!Source)
        txtPONO = Null2String(rsPO!po_no)
        txtDatePO = Null2String(rsPO!DATEORDERED)
        cboModelDescript = Null2String(rsPO!ModelDescript)
        txtModel = Null2String(rsPO!Model)
        txtModelCode = Null2String(rsPO!ModelCode)
        txtModelYear = Null2String(rsPO!MODELYEAR)
        cboColor = Null2String(rsPO!Color)
        txtNotes = Null2String(rsPO!Notes)
        txtFuel = Null2String(rsPO!Fuel)
        txtID = Null2String(rsPO!ID)
        txtDueDate = Null2String(rsPO!DATEREQ)
        txtCusCode = Null2String(rsPO!CUSCDE)
        cboPy_FinLcIssuingBank = Null2String(rsPO!FinLcIssuingBank)
        txtPy_LCNo = Null2String(rsPO!LCNo)
        cboPy_CD_BankName = Null2String(rsPO!CD_BankName)
        txtPy_CD_CheckNo = Null2String(rsPO!CD_CheckNo)
        'txtPY_CD_Date = Null2Date(rsPO!CD_Date)
        txtSubsidy = FormatNumber(NumericVal(rsPO!SUBSIDY))
        txtPy_CD_Amount = FormatNumber(NumericVal(rsPO!CD_AMOUNT))

        cboModeOfPayment = GetModeOfPayment(Null2String(rsPO!modeofpayment))

        If IsDate(rsPO!datereceived) = True Then
            On Error Resume Next
            Set MRRSTATUS = gconDMIS.Execute("SELECT A.STATUS FROM SMIS_MRRINV_TABLE A INNER JOIN SMIS_PO B ON B.PO_NO=A.PONO AND B.PO_NO='" & txtPONO & "'")
            Set MRRSTATUS = gconDMIS.Execute("SELECT Status,pono from SMIS_MRRINV_TABLE where PONO='" & txtPONO & "'")

            If Null2String(MRRSTATUS!STATUS) = "U" Then
                lblSTATUS = "**RECEIVED But Not Posted**"
                cmdCancelCO.Enabled = True
                cmdUnPost.Enabled = False
                cmdPost.Enabled = False
                cmdEdit.Enabled = False
                Exit Sub
            End If

            lblSTATUS = "***RECEIVED***"

            cmdCancelCO.Enabled = False
            cmdUnPost.Enabled = False
            cmdPost.Enabled = False
            cmdEdit.Enabled = False
        Else
            If Null2String(rsPO!STATUS) = "C" Then
                cmdCancelCO.Enabled = False
                cmdUnPost.Enabled = False
                cmdPost.Enabled = False
                lblSTATUS = "***CANCELLED***"
                cmdEdit.Enabled = False
                cmdPrint.Enabled = False
            ElseIf Null2String(rsPO!STATUS) = "P" Then
                cmdCancelCO.Enabled = False
                cmdUnPost.Enabled = True
                cmdPost.Enabled = False
                lblSTATUS = "***POSTED ***"
                cmdEdit.Enabled = False
                cmdPrint.Enabled = True
            Else
                cmdCancelCO.Enabled = True
                cmdUnPost.Enabled = False
                cmdPost.Enabled = True
                lblSTATUS = ""
                cmdEdit.Enabled = True
                cmdPrint.Enabled = False
            End If
        End If

        If NumericVal(rsPO!WTAX) = 0 Then
            Check1.Value = 0
            TXTWTAX = "0.00"
            TXTWTAX.Visible = False
        Else

            Check1.Value = 1
            TXTWTAX = FormatNumber(NumericVal(rsPO!WTAX))
            TXTWTAX.Visible = True
        End If
    Else

        ShowNoRecord
        cmdAdd.Value = True
    End If

End Sub

Private Sub cboModelDescript_Change()
    'If AddorEdit = "ADD" Then
    If cboModelDescript.ListIndex <> -1 And DontChange = False Then
        SetModelLine cboModelDescript, False
    End If
    DontChange = False
    'End If
End Sub

Private Sub cboModelDescript_Click()
    cboModelDescript_Change
End Sub

Private Sub cboModelDescript_Validate(Cancel As Boolean)
    cboModelDescript.ListIndex = SelectCombo(cboModelDescript, cboModelDescript)
    If cboModelDescript.ListIndex = -1 Then
        txtModelCode = ""
        txtModel = ""
        txtPy_CD_Amount = "0.00"
    End If
End Sub

Private Sub cboModeOfPayment_Change()
    If SetModeOfPayment(cboModeOfPayment) = "CA" Then
        fraCrNo.Enabled = False
        '       fraCheckDetail.Enabled = False
    Else
        fraCrNo.Enabled = True
        '     fraCheckDetail.Enabled = True
    End If
End Sub

Private Sub cboModeOfPayment_Click()
    cboModeOfPayment_Change
End Sub

Private Sub cboSource_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        TXTWTAX = 0
        TXTWTAX.Visible = False
    Else
        TXTWTAX.Visible = True
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "PURCHASE ORDER") = False Then Exit Sub
    AddorEdit = "ADD"
    initMemvars
    txtID = 0
    txtPONO = GenerateCode("SMIS_PO", "PO_NO", "000000")
    txtDatePO = Format(LOGDATE, "mm/dd/yyyy")
    picAdds.Visible = False
    picSaves.Visible = True
    fraCrNo.Enabled = True
    txtDatePO.Enabled = False
    fraCheckDetail.Enabled = True
    picTop.Enabled = True
    fraSearch.Enabled = False
    On Error Resume Next
    txtPONO.SetFocus
End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    picAdds.Visible = True
    picSaves.Visible = False
    fraCrNo.Enabled = False
    fraCheckDetail.Enabled = False
    picTop.Enabled = False
    fraSearch.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "PURCHASE ORDER") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox("Do you Want to Cancel this Transaction ", vbOKCancel + vbExclamation, "Confirm Posting") = vbCancel Then Exit Sub
    cmdCancelCO.Enabled = True
    SQL_STATEMENT = ("UPDate SMIS_PO Set Status='C'  Where ID=" & txtID)

    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "C", "PURCHASE ORDER", SQL_STATEMENT, txtID, "", "PO NO:" & txtPONO, "", ""

    'LogAudit "C", "PURCHASE ORDER", cboSource & " PO NO " & txtPONO & " " & cboModelDescript
    rsRefresh
    rsPO.Find ("ID=" & txtID)
    StoreMemVars
    MessagePop RecSaveOk, "Transaction Cancelled", "Record Sucessfully Cancelled", 1000, 2
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdDelTechnicalInquiry_Click()
    Form_KeyDown 116, 1
End Sub

Private Sub cmdEdit_Click()
    'If lblStatus <> "" Then Exit Sub
    If Function_Access(LOGID, "Acess_EDIT", "PURCHASE ORDER") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If NumericVal(txtID) <> 0 Then
        AddorEdit = "EDIT"
        picAdds.Visible = False
        picSaves.Visible = True
        fraCrNo.Enabled = True
        fraCheckDetail.Enabled = True
        txtDatePO.Enabled = False: txtDueDate.Enabled = False
        picTop.Enabled = True
        On Error Resume Next
        cboModelDescript.SetFocus
    End If
    fraSearch.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsPO.MoveNext
    If rsPO.EOF Then
        rsPO.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdok_Click()
    If Function_Access(LOGID, "Acess_PRINT", "PURCHASE ORDER") = False Then Exit Sub
    If LABALLOWREPRINT <> "" Then
        If AllowReprint("PURCHASE ORDER") = False Then Exit Sub
    End If

    On Error GoTo ErrorCode:
    Screen.MousePointer = 11
    LoadSignatories ("PURCHASE ORDER")
    rptPO.Formulas(0) = "ApprovedBy= '" & Null2String(ApprovedBy) & "'"
    rptPO.Formulas(1) = "Preparedby= '" & Null2String(PreparedBy) & "'"
    rptPO.Formulas(2) = "CheckedBy= '" & Null2String(FinancingManager) & "'"
    rptPO.Formulas(3) = "GM='" & Null2String(GeneralManager) & "'"
    rptPO.Formulas(4) = "CompanyName='" & Null2String(COMPANY_NAME) & "'"
    rptPO.Formulas(5) = "CompanyAddress='" & Null2String(COMPANY_ADDRESS) & "'"

    If COMPANY_CODE = "HGC" Then
        rptPO.Formulas(7) = "contactperson='" & txtcontact.Text & "'"
    End If

    If IsDate(txtDatePO) = False Then
        MsgBox "Please Input Valid PO Date", vbInformation
        Exit Sub
    End If
    rptPO.Formulas(6) = "yeer='" & Null2String(Year(txtDatePO)) & "'"
    If N2Str2IntZero(rsPO!WTAX) = 0 Then
        PrintSQLReport rptPO, SMIS_REPORT_PATH & "POWOTAX.rpt", "{SMIS_PO.PO_NO}='" & txtPONO.Text & "'", DMIS_Connection, 1
    Else
        PrintSQLReport rptPO, SMIS_REPORT_PATH & "PO.rpt", "{SMIS_PO.PO_NO}='" & txtPONO.Text & "'", DMIS_Connection, 1
    End If
    LogAudit "V", "PURCHASE ORDER", cboSource & " PO NO " & txtPONO & " " & cboModelDescript

    Screen.MousePointer = 0

    If rptPO.RecordsPrinted = 1 Then
        gconDMIS.Execute ("UPDATE SMIS_PO SET PRINTED=1 WHERE PO_NO='" & txtPONO & "'")
        rsRefresh
        rsPO.Find ("PO_NO='" & txtPONO & "'")
        StoreMemVars
    End If
    picContact.Visible = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdpicCancel_Click()
    picContact.Visible = False
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "PURCHASE ORDER") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox("Do you Want to Post this Transaction ", vbYesNo + vbExclamation, "Confirm Posting") = vbNo Then Exit Sub
    cmdCancelCO.Enabled = False
    SQL_STATEMENT = ("UPDate SMIS_PO  Set Status='P' Where ID=" & txtID)
    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "P", "PURCHASE ORDER", SQL_STATEMENT, txtID, "", "PO No:" & txtPONO, "", ""

    rsRefresh
    rsPO.Find ("ID=" & txtID)
    StoreMemVars
    MessagePop RecSaveOk, "Transaction Posted", "Record Sucessfully Posted", 1000, 2
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()


    rsPO.MovePrevious

    If rsPO.BOF Then
        rsPO.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars


End Sub

Private Sub cmdPrint_Click()
    
    If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HAI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HAS" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HOT" Then
        picContact.Visible = False
    Else
        picContact.Visible = True
    End If
    '        : Update By BTT : 08012008 : Code For HGC
    If Function_Access(LOGID, "Acess_PRINT", "PURCHASE ORDER") = False Then Exit Sub

    If LABALLOWREPRINT <> "" Then
        If AllowReprint("PURCHASE ORDER") = False Then Exit Sub
    End If

    On Error GoTo ErrorCode:
    Screen.MousePointer = 11
    LoadSignatories ("PURCHASE ORDER")
    rptPO.Formulas(0) = "ApprovedBy= '" & Null2String(ApprovedBy) & "'"
    rptPO.Formulas(1) = "Preparedby= '" & Null2String(PreparedBy) & "'"
    rptPO.Formulas(2) = "CheckedBy= '" & Null2String(FinancingManager) & "'"
    rptPO.Formulas(3) = "GM='" & Null2String(GeneralManager) & "'"
    rptPO.Formulas(4) = "CompanyName='" & Null2String(COMPANY_NAME) & "'"
    rptPO.Formulas(5) = "COMPANYADDRESS='" & Null2String(COMPANY_ADDRESS) & "'"
    
    If COMPANY_CODE = "HBK" Then
        rptPO.Formulas(3) = "D_APPBY='" & Null2String(SalesApprovedDesig) & "'"
        rptPO.Formulas(4) = "D_PREPBY='" & Null2String(PreparedByDesig) & "'"
        rptPO.Formulas(5) = "D_CHECKBY='" & Null2String(CheckedByDesig) & "'"
    End If



    If IsDate(txtDatePO) = False Then
        MsgBox "Please Input Valid PO Date", vbInformation
        Exit Sub
    End If
    rptPO.Formulas(6) = "yeer='" & Null2String(Year(txtDatePO)) & "'"
    If N2Str2IntZero(rsPO!WTAX) > 0 Then
        'rptPO.PrinterSelect
        PrintSQLReport rptPO, SMIS_REPORT_PATH & "POWOTAX.rpt", "{SMIS_PO.PO_NO}='" & txtPONO.Text & "'", DMIS_Connection, 1
    Else
        PrintSQLReport rptPO, SMIS_REPORT_PATH & "PO.rpt", "{SMIS_PO.PO_NO}='" & txtPONO.Text & "'", DMIS_Connection, 1
    End If

    NEW_LogAudit "V", "PURCHASE ORDER", "", txtID, "", "PO NO:" & txtPONO, "", ""

    Screen.MousePointer = 0

    If rptPO.RecordsPrinted = 1 Then
        gconDMIS.Execute ("UPDATE SMIS_PO SET PRINTED=1 WHERE PO_NO='" & txtPONO & "'")
        rsRefresh
        rsPO.Find ("PO_NO='" & txtPONO & "'")
        StoreMemVars
    End If


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    Dim lng                                                           As Long
    Dim SQL                                                           As String
    If MULTIPLEPO = True Then
        ShowHidePictureBox2 picmultiple, True
        Exit Sub
    End If
    On Error GoTo ErrorCode:

    If RTrim(LTrim(txtPONO)) = "" Then
        MessagePop RecSaveError, "MISSING FIELDS", "PO NUMBER"
        On Error Resume Next
        txtPONO.SetFocus
        Exit Sub
    End If
    If IsDate(txtDatePO) = False Then
        MessagePop RecSaveError, "Invalid Date", "Date Of PO is Required Field"
        On Error Resume Next
        txtDatePO.SetFocus
        Exit Sub
    End If

    If NumericVal(txtPy_CD_Amount) = 0 Then
        If MsgBox(" Zero Amount ! Are You Sure ?", vbQuestion + vbYesNo) = vbNo Then
            On Error Resume Next
            txtPy_CD_Amount.SetFocus
            Exit Sub
        End If

    End If
    If IsDate(txtDueDate) = False Then
        MessagePop RecSaveError, "Invalid Date", "Date Required is Required Field"
        On Error Resume Next
        txtDueDate.SetFocus
        Exit Sub
    End If

    If Null2String(txtModelCode) = "" Then
        MessagePop RecSaveError, "Invalid Code", "Code is Required Field"
        On Error Resume Next
        txtModelCode.SetFocus
        Exit Sub
    End If
    lng = gconDMIS.Execute("select Count(*) from SMIS_PO WHERE PO_NO=" & N2Str2Null(txtPONO)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Purchase Order Number Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsPO!po_no)) <> UCase(txtPONO) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Sales Order Already Exist"
            Exit Sub
        End If
    End If
    If Null2String(cboModelDescript) = "" Then
        MessagePop RecSaveError, "Invalid Model Description", "Description is Required Field"
        On Error Resume Next
        cboModelDescript.SetFocus
        Exit Sub
    End If

    If Null2String(cboColor) = "" Then
        MessagePop RecSaveError, "Invalid Model Color", "Color is Required Field"
        On Error Resume Next
        cboColor.SetFocus
        Exit Sub
    End If

    If Null2String(cboSource) = "" Then
        On Error Resume Next
        cboSource.SetFocus
        Exit Sub
    End If
    ''''''

    lng = gconDMIS.Execute("select Count(*) from SMIS_PO  WHERE PO_NO=" & N2Str2Null(txtPONO)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "PO Number Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsPO!po_no)) <> UCase(txtPONO) Then
            MessagePop RecSaveWarning, "Duplicate Record", "PO Number Already Exist"
            Exit Sub
        End If
    End If

    Dim temprs                                                        As ADODB.Recordset
    Dim rsHanapID                                                     As ADODB.Recordset
    Dim vID                                                           As String

    Set rsHanapID = New ADODB.Recordset
'Updated: ACL
'Description: Added field Mode of Payment
'Date: 3272011
    If AddorEdit = "ADD" Then

        SQL = " INSERT INTO SMIS_PO "
        SQL = SQL & " ( PO_NO, DateOrdered, ModelDescript"
        SQL = SQL & " , FinLcIssuingBank , LCNo "
        SQL = SQL & " , CD_BankName , CD_CheckNo, CD_Date,CD_Amount "
        SQL = SQL & " , Model, ModelYear, ModelCode,CUSCDE,DateReq, Source, Color, Fuel, Notes, subsidy,WTAX,MODEOFPAYMENT) "
        SQL = SQL & " VALUES( "
        SQL = SQL & N2Str2Null(txtPONO) & "," & N2Str2Null(txtDatePO) & " ," & N2Str2Null(cboModelDescript) & ", "
        SQL = SQL & N2Str2Null(cboPy_FinLcIssuingBank) & "," & N2Str2Null(txtPy_LCNo) & " ,"
        SQL = SQL & N2Str2Null(cboPy_CD_BankName) & "," & N2Str2Null(txtPy_CD_CheckNo) & " ," & N2Str2Null(txPy_CD_Date) & "," & NumericVal(txtPy_CD_Amount) & ","
        SQL = SQL & N2Str2Null(txtModel) & "," & N2Str2Null(txtModelYear) & " ," & N2Str2Null(txtModelCode) & ", "
        SQL = SQL & N2Str2Null(txtCusCode) & "," & N2Date2Null(txtDueDate) & "," & vbCrLf
        SQL = SQL & N2Str2Null(cboSource) & "," & N2Str2Null(cboColor) & " ," & N2Str2Null(txtFuel) & ", " & N2Str2Null(txtNotes) & ", "
        SQL = SQL & NumericVal(txtSubsidy) & "," & NumericVal(TXTWTAX) & "," & N2Str2Null(SetModeOfPayment(cboModeOfPayment)) & ")" & vbCrLf
        SQL = SQL & " SELECT @@IDENTITY "
        Set temprs = gconDMIS.Execute(SQL)

        SQL_STATEMENT = SQL

        NEW_LogAudit "A", "PURCHASE ORDER", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtPONO), "PO_NO", "SMIS_PO"), "", "PO NO:" & txtPONO, "", ""
        ShowSuccessFullyAdded
    Else

        SQL = " Update SMIS_PO "
        SQL = SQL & " SET "
        SQL = SQL & " PO_NO=" & N2Str2Null(txtPONO) & ","
        SQL = SQL & " DateOrdered=" & N2Str2Null(txtDatePO) & ","
        SQL = SQL & " DateReq=" & N2Str2Null(txtDueDate) & ","
        SQL = SQL & " CUSCDE=" & N2Str2Null(txtCusCode) & ","
        SQL = SQL & " ModelDescript=" & N2Str2Null(cboModelDescript) & ","
        SQL = SQL & " Model=" & N2Str2Null(txtModel) & ","
        SQL = SQL & " ModelYear=" & N2Str2Null(txtModelYear) & ","
        SQL = SQL & " Source=" & N2Str2Null(cboSource) & ","
        SQL = SQL & " ModelCode=" & N2Str2Null(txtModelCode) & ","
        SQL = SQL & " Color=" & N2Str2Null(cboColor) & ","
        SQL = SQL & " Fuel=" & N2Str2Null(txtFuel) & ","
        SQL = SQL & " Notes=" & N2Str2Null(txtNotes) & ","
        SQL = SQL & " FinLcIssuingBank=" & N2Str2Null(cboPy_FinLcIssuingBank) & ","
        SQL = SQL & " LCNo=" & N2Str2Null(txtPy_LCNo) & ","
        SQL = SQL & " CD_BankName=" & N2Str2Null(cboPy_CD_BankName) & ","
        SQL = SQL & " CD_CheckNo=" & N2Str2Null(txtPy_CD_CheckNo) & ","
        SQL = SQL & " CD_Date=" & N2Date2Null(txPy_CD_Date) & ","
        SQL = SQL & " CD_Amount=" & NumericVal(txtPy_CD_Amount) & ","
        SQL = SQL & " WTAX=" & NumericVal(TXTWTAX) & ","
        SQL = SQL & " SUBSIDY=" & NumericVal(txtSubsidy) & ","
        SQL = SQL & " ModeOfPayment=" & N2Str2Null(SetModeOfPayment(cboModeOfPayment))
        SQL = SQL & " WHERE ID=" & N2Str2Null(txtID)
        Set temprs = gconDMIS.Execute(SQL)

        '*******************
        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "PURCHASE ORDER", SQL_STATEMENT, Null2String(txtID), "", "PO NO:" & txtPONO, "", ""
        '*******************

        'LogAudit "E", "PURCHASE ORDER", cboSource & " PO NO " & txtPONO & " " & cboModelDescript
        ShowSuccessFullyUpdated
    End If



    Set temprs = temprs.NextRecordset
    If Not temprs Is Nothing Then
        txtID = temprs.Collect(0)
    End If

    picAdds.Visible = True
    picSaves.Visible = False
    InitCombo
    rsRefresh
    rsPO.Find ("ID=" & txtID)
    CboRefresh
    cmdCancel.Value = True

    FillSearchGrid



    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UnPost", "PURCHASE ORDER") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox("Are You Sure You Want to Un-Post this transaction", vbInformation + vbYesNo) = vbNo Then: Exit Sub
    cmdCancelCO.Enabled = True
    SQL_STATEMENT = ("UPDate SMIS_PO  Set Status='U' Where ID=" & txtID)

    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "U", "PURCHASE ORDER", SQL_STATEMENT, txtID, "", "PO NO:" & txtPONO, "", ""


    'LogAudit "U", "PURCHASE ORDER", cboSource & " PO NO " & txtPONO & " " & cboModelDescript
    rsRefresh
    rsPO.Find ("ID=" & txtID)
    StoreMemVars
    MessagePop RecSaveOk, "Transaction Unposted", "Record Sucessfully Un-Posted", 1000, 2
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command1_Click()
    SearchMaster.SearchForCustomers
    SearchMaster.Show 1

End Sub

Private Sub Command2_Click()
    MULTIPLEPO = True: cmdAdd.Value = True

End Sub

Private Sub Command3_Click()
    ShowHidePictureBox2 picmultiple, False, picTop
End Sub

Private Sub Command4_Click()
    On Error GoTo ErrorCode:
    Dim lng                                                           As Long
    Dim i                                                             As Integer
    Dim SQL                                                           As String

    If IsNumeric(txtMultiplePONo) = False Then
        MessagePop RecSaveError, "Invalid Entry", "Invalid Quantity"
        On Error Resume Next
        txtMultiplePONo.SetFocus
        Exit Sub
    End If


    If txtMultiplePONo > 50 Then
        If MsgBox("Are You Sure The Quantity You Have Ordered Is " & txtMultiplePONo & " In Quantity", vbInformation + vbYesNo) = vbNo Then
            On Error Resume Next
            txtMultiplePONo.SetFocus
            Exit Sub
        End If
    End If

    If RTrim(LTrim(txtPONO)) = "" Then
        MessagePop RecSaveError, "MISSING FIELDS", "PO NUMBER"
        On Error Resume Next
        txtPONO.SetFocus
        Exit Sub
    End If


    If IsDate(txtDatePO) = False Then
        MessagePop RecSaveError, "Invalid Date", "Date Of PO is Required Field"
        On Error Resume Next
        txtDatePO.SetFocus
        Exit Sub
    End If

    If NumericVal(txtPy_CD_Amount) = 0 Then
        If MsgBox(" Zero Amount ! Are You Sure ?", vbQuestion + vbYesNo) = vbNo Then
            On Error Resume Next
            txtPy_CD_Amount.SetFocus
            Exit Sub
        End If

    End If


    If IsDate(txtDueDate) = False Then
        MessagePop RecSaveError, "Invalid Date", "Date Required is Required Field"
        On Error Resume Next
        txtDueDate.SetFocus
        Exit Sub
    End If

    If Null2String(txtModelCode) = "" Then
        MessagePop RecSaveError, "Invalid Code", "Code is Required Field"
        On Error Resume Next
        txtModelCode.SetFocus
        Exit Sub
    End If

    ''''''

    lng = gconDMIS.Execute("select Count(*) from SMIS_PO WHERE PO_NO=" & N2Str2Null(txtPONO)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Purchase Order Number Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsPO!po_no)) <> UCase(txtPONO) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Sales Order Already Exist"
            Exit Sub
        End If
    End If

    If Null2String(cboModelDescript) = "" Then
        MessagePop RecSaveError, "Invalid Model Description", "Description is Required Field"
        On Error Resume Next
        cboModelDescript.SetFocus
        Exit Sub
    End If

    If Null2String(cboColor) = "" Then
        MessagePop RecSaveError, "Invalid Model Color", "Color is Required Field"
        On Error Resume Next
        cboColor.SetFocus
        Exit Sub
    End If

    If Null2String(cboSource) = "" Then
        On Error Resume Next
        cboSource.SetFocus
        Exit Sub
    End If
    ''''''

    lng = gconDMIS.Execute("select Count(*) from SMIS_PO  WHERE PO_NO=" & N2Str2Null(txtPONO)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "PO Number Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsPO!po_no)) <> UCase(txtPONO) Then
            MessagePop RecSaveWarning, "Duplicate Record", "PO Number Already Exist"
            Exit Sub
        End If
    End If

    Dim vtxtpono
    Dim temprs                                                        As ADODB.Recordset

    If MsgBox("Are You Sure Quantity You Have Entered Is Correct" & vbCrLf & "Transaction Cannot Be Delete Upon Creation.", vbCritical + vbYesNo) = vbNo Then
        On Error Resume Next
        txtMultiplePONo.SetFocus
        Exit Sub
    End If

    For i = 1 To txtMultiplePONo
        vtxtpono = GenerateCode("SMIS_PO", "PO_NO", "000000")
        SQL = " INSERT INTO SMIS_PO "
        SQL = SQL & " ( PO_NO, DateOrdered, ModelDescript"
        SQL = SQL & " , FinLcIssuingBank , LCNo "
        SQL = SQL & " , CD_BankName , CD_CheckNo, CD_Date,CD_Amount "
        SQL = SQL & " , Model, ModelYear, ModelCode,CUSCDE,DateReq, Source, Color, Fuel, Notes, subsidy) "
        SQL = SQL & " VALUES( "
        SQL = SQL & N2Str2Null(vtxtpono) & "," & N2Str2Null(txtDatePO) & " ," & N2Str2Null(cboModelDescript) & ", "
        SQL = SQL & N2Str2Null(cboPy_FinLcIssuingBank) & "," & N2Str2Null(txtPy_LCNo) & " ,"
        SQL = SQL & N2Str2Null(cboPy_CD_BankName) & "," & N2Str2Null(txtPy_CD_CheckNo) & " ," & N2Str2Null(txPy_CD_Date) & "," & NumericVal(txtPy_CD_Amount) & ","
        SQL = SQL & N2Str2Null(txtModel) & "," & N2Str2Null(txtModelYear) & " ," & N2Str2Null(txtModelCode) & ", "
        SQL = SQL & N2Str2Null(txtCusCode) & "," & N2Date2Null(txtDueDate) & "," & vbCrLf
        SQL = SQL & N2Str2Null(cboSource) & "," & N2Str2Null(cboColor) & " ," & N2Str2Null(txtFuel) & ", " & N2Str2Null(txtNotes) & ", "
        SQL = SQL & NumericVal(txtSubsidy) & ")" & vbCrLf
        SQL = SQL & " SELECT @@IDENTITY "
        Set temprs = gconDMIS.Execute(SQL)
        LogAudit "C", "PURCHASE ORDER MULTIPLE PO", cboSource & " PO NO " & txtPONO & " " & cboModelDescript
    Next



    MULTIPLEPO = False
    ShowHidePictureBox2 picmultiple, False, picTop
    rsRefresh
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command5_Click()
    If Function_Access(LOGID, "ACESS_SYSTEM", "PURCHASE ORDER") = False Then Exit Sub
    txtDatePO.Enabled = True: txtDatePO.SetFocus
End Sub

Private Sub Command6_Click()
    If AddorEdit = "EDIT" Then
        If Function_Access(LOGID, "ACESS_SYSTEM", "PURCHASE ORDER") = False Then Exit Sub
        txtDueDate.Enabled = True: txtDueDate.SetFocus
    End If
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PURCHASE ORDER)"
            'Call frmALL_AuditInquiry.DisplayHistory(labid, "PURCHASE ORDER")
            Call frmALL_AuditInquiry.DisplayHistory(txtID, "PURCHASE ORDER")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    picAdds.Visible = True
    picSaves.Visible = False
    picTop.Enabled = False
    fraCrNo.Enabled = False
    fraCheckDetail.Enabled = False

    InitCombo
    CboRefresh
    Call AddColumnHeader("Date, PO", lstPO)
    Call ResizeColumnHeader(lstPO, "40,50")
    Set SearchMaster = New frmSMIS_Mis_SearchMaster
    picContact.Visible = False
    rsRefresh
    If Not rsPO.EOF And Not rsPO.BOF Then
        rsPO.MoveLast
    End If
    initMemvars
    StoreMemVars
End Sub

Private Sub lstPO_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPO
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

Private Sub lstPO_DblClick()
    If lstPO.SelectedItem Is Nothing Then: Exit Sub
    If cmdEdit.Enabled = False Then Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub lstPO_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ADDER:
    rsPO.MoveFirst
    rsPO.Find ("ID=" & Item.ListSubItems(2).Text)
    StoreMemVars
    Exit Sub
ADDER:
    Err.Clear
End Sub

Private Sub optDate_Click()
    textSearch_Change
End Sub

Private Sub optPO_Click()
    textSearch_Change
End Sub

Private Sub optVModel_Click()
    textSearch_Change
End Sub

Private Sub SearchMaster_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    txtCusCode = oCusRs!CUSCDE
    Unload SearchMaster
End Sub

Private Sub textSearch_Change()
    FillSearchGrid
End Sub

Private Sub Timer1_Timer()
    If lblSTATUS.Caption <> "" Then
        If lblSTATUS.Visible = True Then
            lblSTATUS.Visible = False
        Else
            lblSTATUS.Visible = True
        End If
    End If
End Sub

Private Sub txtACC_QTY_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtACC_SRP_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtACC_Total_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtDatePO_GotFocus()
    If IsDate(txtDatePO) = False Then
        txtDatePO = ""
    Else
        txtDatePO = Format(txtDatePO, "mm/dd/yyyy")
    End If
End Sub

Private Sub txtDatePO_LostFocus()
    If IsDate(txtDatePO) = False Then
        txtDatePO = ""
    Else
        txtDatePO = Format(txtDatePO, "mmm dd yyyy")
        txtDueDate = Format(txtDatePO, "mmm dd yyyy")
    End If
End Sub

Private Sub txtDueDate_GotFocus()
    If IsDate(txtDueDate) = False Then
        txtDueDate = ""
    Else
        txtDueDate = Format(txtDueDate, "mm/dd/yyyy")
    End If

End Sub

Private Sub txtDueDate_LostFocus()
    If IsDate(txtDueDate) = False Then
        txtDueDate = ""
    Else
        txtDueDate = Format(txtDueDate, "mmm dd yyyy")
    End If
End Sub

Private Sub txtFuel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)

End Sub

Private Sub txtModelCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If txtModelCode <> "" And KeyCode = 13 Then
        If AddorEdit = "add" Then
            SetModelLine txtModelCode, True
        End If
    ElseIf KeyCode = vbKeyEscape And AddorEdit = "EDIT" Then
        DontChange = True
        txtModelCode = Null2String(rsPO!ModelCode)
        txtModel = Null2String(rsPO!Model)
        txtModelCode = Null2String(rsPO!ModelCode)

    End If
End Sub

Private Sub txtModelCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtModelCode_LostFocus()

    DontChange = False
End Sub

Private Sub txtModelYear_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtMultiplePONo_GotFocus()
    If NumericVal(txtMultiplePONo.Text) <= 0 Then txtMultiplePONo = ""
End Sub

Private Sub txtMultiplePONo_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtMultiplePONo_LostFocus()
    If NumericVal(txtMultiplePONo) <= 0 Then txtMultiplePONo = "0.00"
    txtMultiplePONo = FormatNumber(NumericVal(txtMultiplePONo))
End Sub

Private Sub txtPONO_LostFocus()
    txtPONO = Format(txtPONO, "000000")
End Sub

Private Sub txtPy_CD_Amount_GotFocus()
    If NumericVal(txtPy_CD_Amount.Text) <= 0 Then
        txtPy_CD_Amount = ""
    Else
        txtPy_CD_Amount = FormatNumber(txtPy_CD_Amount)
    End If

End Sub

Private Sub txtPy_CD_Amount_LostFocus()
    If NumericVal(txtPy_CD_Amount) <= 0 Then txtPy_CD_Amount = "0.00"
    txtPy_CD_Amount = FormatNumber(NumericVal(txtPy_CD_Amount))
End Sub

Private Sub txtSubsidy_GotFocus()
    If NumericVal(txtSubsidy.Text) <= 0 Then
        txtSubsidy = ""
    Else
        txtSubsidy = FormatNumber(txtSubsidy)
    End If
End Sub

Private Sub txtSubsidy_LostFocus()
    If NumericVal(txtSubsidy) <= 0 Then txtSubsidy = "0.00"
    txtSubsidy = FormatNumber(NumericVal(txtSubsidy))
End Sub

Private Sub TXTWTAX_GotFocus()
    If NumericVal(TXTWTAX.Text) <= 0 Then
        TXTWTAX = "0.00"
    Else
        TXTWTAX = FormatNumber(TXTWTAX)
    End If

End Sub

Private Sub TXTWTAX_LostFocus()
    If NumericVal(TXTWTAX.Text) <= 0 Then TXTWTAX = ""
End Sub

