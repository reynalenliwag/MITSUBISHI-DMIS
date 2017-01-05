VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frm_TOOLS_ARMONITORING 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Recievable "
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   14160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   14160
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   4290
      ScaleHeight     =   945
      ScaleWidth      =   5625
      TabIndex        =   61
      Top             =   3840
      Width           =   5625
      Begin VB.PictureBox Picture5 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   5805
         TabIndex        =   62
         Top             =   0
         Width           =   5805
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   30
            TabIndex        =   63
            Top             =   0
            Width           =   3105
         End
      End
      Begin MSComctlLib.ProgressBar PROGBAR 
         Height          =   405
         Left            =   30
         TabIndex        =   64
         Top             =   480
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label22 
         Caption         =   "Label20"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   65
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.PictureBox picAllLedger 
      Height          =   6255
      Left            =   480
      ScaleHeight     =   6195
      ScaleWidth      =   13305
      TabIndex        =   48
      Top             =   840
      Width           =   13365
      Begin VB.PictureBox Picture3 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   13275
         TabIndex        =   53
         Top             =   0
         Width           =   13275
         Begin VB.CommandButton Command2 
            Caption         =   "X"
            Height          =   195
            Left            =   12960
            TabIndex        =   55
            Top             =   30
            Width           =   255
         End
         Begin VB.Label Label19 
            BackColor       =   &H8000000D&
            Caption         =   "Customer ledger:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   30
            TabIndex        =   57
            Top             =   0
            Width           =   1785
         End
         Begin VB.Label lblledger 
            BackColor       =   &H8000000D&
            Caption         =   "Customer ledger:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   285
            Left            =   1800
            TabIndex        =   54
            Top             =   0
            Width           =   5745
         End
      End
      Begin VB.TextBox txtcustomerledger 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   9840
         TabIndex        =   51
         Top             =   5820
         Width           =   2445
      End
      Begin MSFlexGridLib.MSFlexGrid grdlAllledger 
         Height          =   5055
         Left            =   30
         TabIndex        =   50
         Top             =   750
         Width           =   13155
         _ExtentX        =   23204
         _ExtentY        =   8916
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorSel    =   12648384
         SelectionMode   =   1
         AllowUserResizing=   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblcustcode 
         Caption         =   "Label21"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   150
         TabIndex        =   59
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label lblcustname 
         Caption         =   "Label21"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   1230
         TabIndex        =   58
         Top             =   360
         Width           =   9555
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Balance:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7620
         TabIndex        =   52
         Top             =   5850
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "CRJ Payment"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6675
      Left            =   3210
      TabIndex        =   25
      Top             =   2280
      Width           =   10935
      Begin VB.TextBox txtrunbal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   8370
         TabIndex        =   47
         Top             =   5430
         Width           =   2445
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   90
         ScaleHeight     =   585
         ScaleWidth      =   10755
         TabIndex        =   35
         Top             =   5790
         Width           =   10785
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Left            =   6780
            TabIndex        =   46
            Text            =   "0.00"
            Top             =   240
            Width           =   1635
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Left            =   4410
            TabIndex        =   45
            Text            =   "0.00"
            Top             =   240
            Width           =   1635
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Left            =   2220
            TabIndex        =   44
            Text            =   "0.00"
            Top             =   240
            Width           =   1635
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Left            =   0
            TabIndex        =   43
            Text            =   "0.00"
            Top             =   240
            Width           =   1635
         End
         Begin VB.TextBox txtbalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Left            =   8460
            TabIndex        =   42
            Text            =   "0"
            Top             =   240
            Width           =   2175
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   10725
            TabIndex        =   36
            Top             =   -30
            Width           =   10725
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               Caption         =   "Total Owning"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   8820
               TabIndex        =   41
               Top             =   30
               Width           =   1755
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               Caption         =   "91 days and more"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6660
               TabIndex        =   40
               Top             =   30
               Width           =   1755
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               Caption         =   "61 to  90 days"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   4260
               TabIndex        =   39
               Top             =   30
               Width           =   1755
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               Caption         =   "31 to  60 days"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2100
               TabIndex        =   38
               Top             =   30
               Width           =   1755
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               Caption         =   "1 to  30 days"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   -510
               TabIndex        =   37
               Top             =   30
               Width           =   2115
            End
         End
      End
      Begin VB.TextBox txtTotalPayment 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   8370
         TabIndex        =   33
         Top             =   5430
         Width           =   2445
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   5175
         Left            =   90
         TabIndex        =   26
         Top             =   270
         Width           =   10785
         _Version        =   655364
         _ExtentX        =   19024
         _ExtentY        =   9128
         _StockProps     =   64
         Color           =   8
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         PaintManager.MinTabWidth=   100
         ItemCount       =   2
         Item(0).Caption =   "Payment"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "GrdCRJPayment"
         Item(1).Caption =   "Ledger per voucher"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "grdcusledger"
         Begin MSFlexGridLib.MSFlexGrid GrdCRJPayment 
            Height          =   4485
            Left            =   30
            TabIndex        =   27
            Top             =   600
            Width           =   10725
            _ExtentX        =   18918
            _ExtentY        =   7911
            _Version        =   393216
            Cols            =   7
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid grdcusledger 
            Height          =   4485
            Left            =   -69970
            TabIndex        =   49
            Top             =   600
            Visible         =   0   'False
            Width           =   10725
            _ExtentX        =   18918
            _ExtentY        =   7911
            _Version        =   393216
            Cols            =   7
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Label Label22 
         Caption         =   "Press F1 to show AR Customer ledger "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   56
         Top             =   6420
         Width           =   3825
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Payment:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   32
         Top             =   5460
         Width           =   2175
      End
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
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
      Left            =   6030
      TabIndex        =   5
      Top             =   600
      Width           =   4185
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sales Journal Impormation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   3210
      TabIndex        =   3
      Top             =   60
      Width           =   10935
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Left            =   4920
         TabIndex        =   60
         Top             =   120
         Width           =   435
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
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
         Left            =   1650
         TabIndex        =   28
         Top             =   540
         Width           =   1155
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5190
         TabIndex        =   24
         Top             =   1170
         Width           =   5655
      End
      Begin VB.TextBox txtReferenceno 
         Appearance      =   0  'Flat
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
         Left            =   8940
         TabIndex        =   22
         Top             =   810
         Width           =   1905
      End
      Begin VB.TextBox txtTerm 
         Appearance      =   0  'Flat
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
         Left            =   5190
         TabIndex        =   20
         Top             =   870
         Width           =   1845
      End
      Begin VB.TextBox txtInvoiceAmount 
         Appearance      =   0  'Flat
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
         Left            =   1650
         TabIndex        =   19
         Top             =   1770
         Width           =   1875
      End
      Begin VB.TextBox txtInvoicedate 
         Appearance      =   0  'Flat
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
         Left            =   1650
         TabIndex        =   17
         Top             =   1170
         Width           =   1875
      End
      Begin VB.TextBox txtRefdate 
         Appearance      =   0  'Flat
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
         Left            =   8940
         TabIndex        =   15
         Top             =   480
         Width           =   1905
      End
      Begin VB.TextBox txtInvoiceno 
         Appearance      =   0  'Flat
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
         Left            =   1650
         TabIndex        =   13
         Top             =   1470
         Width           =   1875
      End
      Begin VB.TextBox txtJdate 
         Appearance      =   0  'Flat
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
         Left            =   8940
         TabIndex        =   11
         Top             =   180
         Width           =   1905
      End
      Begin VB.TextBox txtInvoicetype 
         Appearance      =   0  'Flat
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
         Left            =   1650
         TabIndex        =   9
         Top             =   870
         Width           =   1875
      End
      Begin VB.TextBox txtVoucher 
         Appearance      =   0  'Flat
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
         Left            =   1650
         TabIndex        =   7
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Ref. No:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7350
         TabIndex        =   29
         Top             =   810
         Width           =   1515
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Particulars:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3630
         TabIndex        =   23
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Term:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3660
         TabIndex        =   21
         Top             =   900
         Width           =   1515
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Invoice Amt:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   18
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Invoice Date:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Ref. Date:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7350
         TabIndex        =   14
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Invoice No:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1470
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Journal Date:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7350
         TabIndex        =   10
         Top             =   210
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Invoice type:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   900
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Voucher No:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Customer Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search By"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3195
      Begin VB.ComboBox cboAccount 
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
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   210
         Width           =   3075
      End
      Begin VB.OptionButton Option2 
         Caption         =   "By Customer Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   31
         Top             =   810
         Width           =   2325
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Voucher no"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   570
         Value           =   -1  'True
         Width           =   2325
      End
      Begin VB.TextBox txtSearch 
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
         Left            =   60
         TabIndex        =   2
         Top             =   1050
         Width           =   3075
      End
      Begin MSComctlLib.ListView LstSJ 
         Height          =   7455
         Left            =   60
         TabIndex        =   1
         Top             =   1350
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   13150
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Voucher"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_TOOLS_ARMONITORING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Thecustcode                                   As String
Private Sub TabStrip2_Change()

End Sub
Sub GetSalesJournal(XXX As String)
    Dim tmp                                       As String
    Dim Item                                      As ListItem
    tmp = XXX
    Dim rsSalesJournal                            As New ADODB.Recordset
    LstSJ.ListItems.Clear
    If tmp <> "" Then
        If Option1.Value = True Then
            'Set rsSalesJournal = gconDMIS.Execute("SELECT  distinct (voucherno),ID,acctname FROM amis_vw_araging where voucherno like '" & XXX & "%' and jtype = 'SJ' and acct_code='" & ReturnAcctCode(cboAccount.Text) & "'  order by voucherno asc")
            Set rsSalesJournal = gconDMIS.Execute("SELECT DISTINCT dbo.AMIS_Journal_HD.VoucherNo AS voucherno, dbo.AMIS_Journal_HD.JType, dbo.AMIS_Journal_HD.ID, dbo.AMIS_Journal_Det.Acct_Code, dbo.ALL_Customer_Table.ACCTNAME FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer_Table ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer_Table.CUSCDE where dbo.AMIS_Journal_HD.VoucherNo like '" & XXX & "%' and (dbo.AMIS_Journal_HD.JType = 'SJ' or dbo.AMIS_Journal_HD.JType = 'COB') and acct_code='" & returnAcctCode(cboAccount.Text) & "' order by voucherno asc")
        Else
            'Set rsSalesJournal = gconDMIS.Execute("SELECT  distinct (voucherno),ID,acctname FROM amis_vw_araging where acctname like '" & XXX & "%' and jtype = 'SJ' and acct_code='" & ReturnAcctCode(cboAccount.Text) & "'  order by voucherno asc")
            Set rsSalesJournal = gconDMIS.Execute("SELECT DISTINCT dbo.AMIS_Journal_HD.VoucherNo AS voucherno, dbo.AMIS_Journal_HD.JType, dbo.AMIS_Journal_HD.ID, dbo.AMIS_Journal_Det.Acct_Code, dbo.ALL_Customer_Table.ACCTNAME FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer_Table ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer_Table.CUSCDE where dbo.ALL_Customer_Table.ACCTNAME like '" & XXX & "%' and (dbo.AMIS_Journal_HD.JType = 'SJ' or dbo.AMIS_Journal_HD.JType = 'COB') and acct_code='" & returnAcctCode(cboAccount.Text) & "' order by voucherno asc")
        End If
    Else
        Set rsSalesJournal = gconDMIS.Execute("SELECT DISTINCT dbo.AMIS_Journal_HD.VoucherNo AS voucherno, dbo.AMIS_Journal_HD.JType, dbo.AMIS_Journal_HD.ID, dbo.AMIS_Journal_Det.Acct_Code, dbo.ALL_Customer_Table.ACCTNAME FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer_Table ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer_Table.CUSCDE where (dbo.AMIS_Journal_HD.JType = 'SJ' or dbo.AMIS_Journal_HD.JType = 'COB') and acct_code='" & returnAcctCode(cboAccount.Text) & "' order by voucherno asc")
    End If
    Do While Not rsSalesJournal.EOF
        Set Item = LstSJ.ListItems.Add(, , rsSalesJournal!ID)
        Item.SubItems(1) = Null2String(rsSalesJournal!VOUCHERNO)
        Item.SubItems(2) = Null2String(rsSalesJournal!AcctName)
        rsSalesJournal.MoveNext
    Loop
    Set rsSalesJournal = Nothing
End Sub

Private Sub cboAccount_Change()
    returnAcctCode cboAccount.Text

End Sub

Private Sub cboAccount_Click()
    returnAcctCode cboAccount.Text
    GetSalesJournal txtSearch.Text
End Sub

Private Sub cboAccount_LostFocus()
    returnAcctCode cboAccount.Text

End Sub



Private Sub Command1_Click()
'processAR
    ar
End Sub

Private Sub Command2_Click()
    picAllLedger.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyF1:
        picAllLedger.Visible = True
        lblledger.Caption = cboAccount.Text
        lblcustname.Caption = txtName.Text
        lblcustcode.Caption = txtCode + ":"
        fillALLCustomerledger txtCode, txtInvoiceNo, txtInvoiceType

    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    InitSJ
    InitCRJPayment
    GetAccountCode
    InitGridLedger
    initAllLedger
    picAllLedger.Visible = False
    Picture4.Visible = False
End Sub

Private Sub GrdCRJPayment_DblClick()
    Dim VARVOUCHERNO
    On Error Resume Next
    GrdCRJPayment.Row = GrdCRJPayment.Row
    GrdCRJPayment.Col = 1
    VARVOUCHERNO = Left(GrdCRJPayment.Text, 6)
    JOURNALTYPE = "CRJ"
    If JOURNALTYPE = "CRJ" Then
        Unload frmAMISJournalEntry
        frmAMISJournalEntry.Show
        frmAMISJournalEntry.StoreSearch (VARVOUCHERNO)
    End If
End Sub

Private Sub LstSJ_ItemClick(ByVal Item As MSComctlLib.ListItem)
    StoreSJ LstSJ.SelectedItem.ListSubItems(1)
    FillGridCRJPayment txtInvoiceNo, txtInvoiceType
    filCustomerLedger txtVoucher, txtInvoiceNo, txtInvoiceType
End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If TabControl1.SelectedItem = 0 Then
        txtrunbal.Visible = False
        txtTotalPayment.Visible = True
        Label12.Caption = "Total Payment:"
    End If
    If TabControl1.SelectedItem = 1 Then
        Label12.Caption = "Balance:"
        txtrunbal.Visible = True
        txtTotalPayment.Visible = False
    End If
End Sub

Private Sub txtSEARCH_Change()
    GetSalesJournal txtSearch.Text
End Sub
Sub InitSJ()
    txtVoucher.Text = ""
    txtName.Text = ""
    txtInvoiceAmount.Text = ""
    txtInvoiceDate.Text = ""
    txtInvoiceType.Text = ""
    txtRemarks.Text = ""
    txtRefDate.Text = ""
    txtInvoiceNo.Text = ""
End Sub
Sub StoreSJ(XXX As String)
    Dim RS                                        As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT dbo.AMIS_Journal_HD.VoucherNo AS voucherno, dbo.AMIS_Journal_HD.JType,dbo.AMIS_Journal_HD.status,dbo.AMIS_Journal_HD.remarks,dbo.AMIS_Journal_HD.debit,dbo.AMIS_Journal_HD.invoiceno,dbo.AMIS_Journal_HD.jdate,dbo.AMIS_Journal_HD.invoicedate,dbo.AMIS_Journal_HD.invoicetype,dbo.AMIS_Journal_HD.customercode,dbo.AMIS_Journal_HD.ID, dbo.AMIS_Journal_Det.Acct_Code, dbo.ALL_Customer_Table.ACCTNAME FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer_Table ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer_Table.CUSCDE where dbo.AMIS_Journal_HD.voucherno = '" & XXX & "'and (dbo.AMIS_Journal_HD.JType = 'SJ' or dbo.AMIS_Journal_HD.JType = 'COB')  and acct_code='" & returnAcctCode(cboAccount.Text) & "' and dbo.AMIS_Journal_HD.status = 'P' order by voucherno asc")
    'Set rs = gconDMIS.Execute("Select * from AMIS_vw_araging where voucherno='" & XXX & "' and jtype = 'SJ' and acct_code='" & ReturnAcctCode(cboAccount.Text) & "'")
    If Not RS.EOF And Not RS.BOF Then
        txtVoucher.Text = Null2String(RS!VOUCHERNO)
        txtCode = (RS!CustomerCode)
        txtName.Text = GetCustomerName(RS!CustomerCode)
        'txtInvoiceAmount.Text = ToDoubleNumber(rs!invoiceamt)
        txtInvoiceAmount.Text = ToDoubleNumber(RS!DEBIT)
        txtInvoiceDate.Text = Null2String(RS!invoicedate)
        txtInvoiceType.Text = Null2String(RS!InvoiceType)
        txtRemarks.Text = Null2String(RS!remarks)
        'txtRefdate.Text = Null2String(rs!refdate)
        txtInvoiceNo.Text = Null2String(RS!INVOICENO)
        txtJDate.Text = Null2String(RS!JDate)
        'txtTerm = Null2String(rs!paytype)
        ' txtReferenceno = Null2String(rs!refno)
    End If
End Sub
Function GetCustomerName(XXX As String)
    Dim RS                                        As New ADODB.Recordset
    Set RS = gconDMIS.Execute("Select ACCTNAME from ALL_customer_table where cuscde='" & XXX & "'")
    If Not RS.EOF And Not RS.BOF Then
        GetCustomerName = Null2String(RS!AcctName)
    End If
End Function
Sub InitCRJPayment()
    With GrdCRJPayment
        .ColWidth(0) = 1200
        .ColWidth(1) = 1500
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 3000
        .Row = 0
        .Col = 0: .Text = "Journal Date "
        .Col = 1: .Text = "Voucher No"
        .Col = 2: .Text = "Type"
        .Col = 3: .Text = "Invoice Date"
        .Col = 4: .Text = "Invoice Type"
        .Col = 5: .Text = "Invoice No"
        .Col = 6: .Text = "Amount"
    End With
End Sub
Sub FillGridCRJPayment(xINVOICENO As String, xINVOICETYPE As String)
    Dim RS                                        As New ADODB.Recordset
    Dim totalpayment                              As Double
    Dim cnt                                       As Integer
    totalpayment = 0
    Dim test                                      As Date
    test = "12/30/2008"
    RS.Open "Select * from amis_crjdetail_total where invoiceno='" & xINVOICENO & "' and invoicetype='" & xINVOICETYPE & "' and acct_Code='" & returnAcctCode(cboAccount.Text) & "' and jdate <= ('" & test & "') and customercode = '" & txtCode.Text & "' ", gconDMIS, adOpenForwardOnly, adLockReadOnly
    cleargrid GrdCRJPayment: InitCRJPayment
    cnt = 0
    If Not RS.EOF And Not RS.BOF Then
        RS.MoveFirst
        Do While Not RS.EOF
            cnt = cnt + 1
            GrdCRJPayment.AddItem ReturnCRJHeader(RS!VOUCHERNO) & Chr(9) & (RS!VOUCHERNO) & Chr(9) & _
                                  (RS!CR_type) & Chr(9) & (RS!invoicedate) & Chr(9) & _
                                  (RS!InvoiceType) & Chr(9) & (RS!INVOICENO) & Chr(9) & _
                                  ToDoubleNumber(RS!invoiceamount)
            totalpayment = totalpayment + NumericVal(RS!invoiceamount)
            RS.MoveNext
        Loop
        If cnt > 0 Then GrdCRJPayment.RemoveItem 1
    End If
    txtTotalPayment = ToDoubleNumber(totalpayment)
    'txtbalance = ToDoubleNumber(NumericVal(txtInvoiceAmount) - NumericVal(txtTotalPayment))
    Set RS = Nothing
End Sub
Function ReturnCRJHeader(XXX As String)
    Dim RS                                        As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT Jdate from AMIS_journal_hd where jtype='CRJ' and voucherno='" & XXX & "'")
    If Not RS.EOF And Not RS.BOF Then
        ReturnCRJHeader = Null2String(RS!JDate)
    End If
    Set RS = Nothing
End Function
Sub GetAccountCode()
    Dim RS                                        As New ADODB.Recordset
    Set RS = gconDMIS.Execute("select acctcode,description from amis_chartaccount where LEFT(acctcode,5)='11-02' order by description asc")
    cboAccount.Clear
    If Not RS.EOF And Not RS.BOF Then
        RS.MoveFirst
        Do While Not RS.EOF
            cboAccount.AddItem Null2String(RS!Description)
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
End Sub
Function returnAcctCode(XXX As String)
    Dim RS                                        As New ADODB.Recordset
    Set RS = gconDMIS.Execute("Select acctcode from Amis_chartAccount where description='" & XXX & "'")
    If Not RS.EOF And Not RS.BOF Then
        returnAcctCode = Null2String(RS!ACCTCODE)
    End If
    Set RS = Nothing
End Function
Sub InitGridLedger()
    With grdcusledger
        .ColWidth(0) = 1200
        .ColWidth(1) = 1500
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 3000
        .Row = 0
        .Col = 0: .Text = "Journal Date "
        .Col = 1: .Text = "Voucher No"
        .Col = 2: .Text = "Type"
        .Col = 3: .Text = "Invoice Date"
        .Col = 4: .Text = "Invoice Type"
        .Col = 5: .Text = "Invoice No"
        .Col = 6: .Text = "Amount"
    End With
End Sub
Sub filCustomerLedger(xVOUCHERNO As String, xINVOICENO As String, xINVOICETYPE As String)
    Dim RsJournalHeader                           As New ADODB.Recordset
    Dim RsCrjDetail                               As New ADODB.Recordset
    Dim cnt                                       As Integer
    Dim BALANCE                                   As Double
    Dim AMOUNT2PAY                                As Double
    cleargrid grdcusledger: InitCRJPayment
    Set RsJournalHeader = gconDMIS.Execute("Select * from AMIS_vw_araging where voucherno='" & xVOUCHERNO & "' and (jtype = 'SJ' or jtype = 'COB')  and acct_code='" & returnAcctCode(cboAccount.Text) & "'")
    If Not RsJournalHeader.EOF And Not RsJournalHeader.BOF Then
        RsJournalHeader.MoveFirst
        Do While Not RsJournalHeader.EOF
            cnt = cnt + 1
            If RsJournalHeader!jtype = "COB" Then
                grdcusledger.AddItem (RsJournalHeader!JDate) & Chr(9) & (RsJournalHeader!VOUCHERNO) & Chr(9) & _
                                     (RsJournalHeader!jtype) & Chr(9) & (RsJournalHeader!invoicedate) & Chr(9) & _
                                     (RsJournalHeader!InvoiceType) & Chr(9) & (RsJournalHeader!INVOICENO) & Chr(9) & _
                                     ToDoubleNumber(RsJournalHeader!InvoiceAmt)
                AMOUNT2PAY = NumericVal(RsJournalHeader!InvoiceAmt)
            Else
                grdcusledger.AddItem (RsJournalHeader!JDate) & Chr(9) & (RsJournalHeader!VOUCHERNO) & Chr(9) & _
                                     (RsJournalHeader!jtype) & Chr(9) & (RsJournalHeader!invoicedate) & Chr(9) & _
                                     (RsJournalHeader!InvoiceType) & Chr(9) & (RsJournalHeader!INVOICENO) & Chr(9) & _
                                     ToDoubleNumber(RsJournalHeader!DEBIT)
                AMOUNT2PAY = NumericVal(RsJournalHeader!DEBIT)
            End If

            'Balance = NumericVal(RsJournalHeader!DEBIT)
            BALANCE = AMOUNT2PAY
            RsJournalHeader.MoveNext

            'Set RsCrjDetail = gconDMIS.Execute("Select * from amis_crjdetail_total where invoiceno='" & xInvoiceNo & "' and invoicetype='" & xInvoiceType & "' and j_class='" & ReturnAcctCode(cboAccount.Text) & "'")
            'Set RsCrjDetail = gconDMIS.Execute("Select * from amis_crjdetail_total where invoiceno='" & xInvoiceNo & "' and invoicetype='" & xInvoiceType & "' and customercode='" & txtCode.Text & "'")
            Set RsCrjDetail = gconDMIS.Execute("Select * from amis_crj_detail where invoiceno='" & xINVOICENO & "' and invoicetype='" & xINVOICETYPE & "'")
            Do While Not RsCrjDetail.EOF
                BALANCE = BALANCE - (RsCrjDetail!invoiceamount)
                grdcusledger.AddItem ReturnCRJHeader(RsCrjDetail!VOUCHERNO) & Chr(9) & (RsCrjDetail!VOUCHERNO) & Chr(9) & _
                                     (RsCrjDetail!CR_type) & Chr(9) & (RsCrjDetail!invoicedate) & Chr(9) & _
                                     (RsCrjDetail!InvoiceType) & Chr(9) & (RsCrjDetail!INVOICENO) & Chr(9) & _
                                     ToDoubleNumber(NumericVal(RsCrjDetail!invoiceamount))

                RsCrjDetail.MoveNext
            Loop
        Loop
        If cnt > 0 Then grdcusledger.RemoveItem 1
        txtrunbal = ToDoubleNumber(NumericVal(BALANCE))
        txtbalance = txtrunbal
    End If
    Set RsJournalHeader = Nothing
End Sub
Sub initAllLedger()
    With grdlAllledger
        .ColWidth(0) = 1200
        .ColWidth(1) = 1500
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 2500
        .ColWidth(7) = 2500
        .Row = 0
        .Col = 0: .Text = "Journal Date "
        .Col = 1: .Text = "Voucher No"
        .Col = 2: .Text = "Type"
        .Col = 3: .Text = "Invoice Date"
        .Col = 4: .Text = "Invoice Type"
        .Col = 5: .Text = "Invoice No"
        .Col = 6: .Text = "Amount"
        .Col = 7: .Text = "Balance":
    End With
End Sub
Sub fillALLCustomerledger(xtheCode As String, xINVOICENO As String, xINVOICETYPE As String)
    Dim RsJournalHeader                           As New ADODB.Recordset
    Dim RsCrjDetail                               As New ADODB.Recordset
    Dim cnt                                       As Integer
    Dim TotalSJ                                   As Double
    Dim TotalCRJ                                  As Double
    Dim AMOUNT2PAY                                As Double
    Dim BALANCE                                   As Double
    Dim BA
    TotalSJ = 0
    TotalCRJ = 0
    cleargrid grdlAllledger: InitCRJPayment
    Set RsJournalHeader = gconDMIS.Execute("Select * from AMIS_vw_araging where customercode='" & xtheCode & "' and (jtype = 'SJ' or jtype = 'COB') and acct_code='" & returnAcctCode(cboAccount.Text) & "' order by jdate asc")
    If Not RsJournalHeader.EOF And Not RsJournalHeader.BOF Then
        RsJournalHeader.MoveFirst
        Do While Not RsJournalHeader.EOF
            cnt = cnt + 1
            If RsJournalHeader!jtype = "COB" Then
                BALANCE = BALANCE + NumericVal(RsJournalHeader!InvoiceAmt)
                grdlAllledger.AddItem (RsJournalHeader!JDate) & Chr(9) & (RsJournalHeader!VOUCHERNO) & Chr(9) & _
                                      (RsJournalHeader!jtype) & Chr(9) & (RsJournalHeader!invoicedate) & Chr(9) & _
                                      (RsJournalHeader!InvoiceType) & Chr(9) & (RsJournalHeader!INVOICENO) & Chr(9) & _
                                      ToDoubleNumber(RsJournalHeader!InvoiceAmt) & Chr(9) & ToDoubleNumber(N2Str2Zero(BALANCE))

            Else
                BALANCE = BALANCE + NumericVal(RsJournalHeader!DEBIT)
                grdlAllledger.AddItem (RsJournalHeader!JDate) & Chr(9) & (RsJournalHeader!VOUCHERNO) & Chr(9) & _
                                      (RsJournalHeader!jtype) & Chr(9) & (RsJournalHeader!invoicedate) & Chr(9) & _
                                      (RsJournalHeader!InvoiceType) & Chr(9) & (RsJournalHeader!INVOICENO) & Chr(9) & _
                                      ToDoubleNumber(RsJournalHeader!DEBIT) & Chr(9) & ToDoubleNumber(N2Str2Zero(BALANCE))

            End If

            Set RsCrjDetail = gconDMIS.Execute("Select * from amis_crj_detail where invoiceno='" & (RsJournalHeader!INVOICENO) & "' and invoicetype='" & (RsJournalHeader!InvoiceType) & "'")
            Do While Not RsCrjDetail.EOF
                'If CheckCRJCustomer(RsCrjDetail!voucherno, xtheCode) = False Then
                BALANCE = BALANCE - NumericVal(RsCrjDetail!invoiceamount)
                grdlAllledger.AddItem ReturnCRJHeader(RsCrjDetail!VOUCHERNO) & Chr(9) & (RsCrjDetail!VOUCHERNO) & Chr(9) & _
                                      (RsCrjDetail!CR_type) & Chr(9) & (RsCrjDetail!invoicedate) & Chr(9) & _
                                      (RsCrjDetail!InvoiceType) & Chr(9) & (RsCrjDetail!INVOICENO) & Chr(9) & _
                                      ToDoubleNumber(NumericVal(RsCrjDetail!invoiceamount)) & Chr(9) & ToDoubleNumber(N2Str2Zero(BALANCE))
                TotalCRJ = TotalCRJ + (RsCrjDetail!invoiceamount)
                RsCrjDetail.MoveNext
                'End If
            Loop
            RsJournalHeader.MoveNext
        Loop
        If cnt > 0 Then grdlAllledger.RemoveItem 1
        txtcustomerledger.Text = ToDoubleNumber(BALANCE)
    End If
    Set RsJournalHeader = Nothing
End Sub
Sub processAR()
    Dim RsJournalHeader                           As New ADODB.Recordset
    Dim RsCrjDetail                               As New ADODB.Recordset
    Dim RSSJVoucherno                             As New ADODB.Recordset
    Dim cnt                                       As Integer
    Dim BALANCE                                   As Double
    Dim test                                      As String
    Dim totalpayment                              As Double
    Dim totalbalance                              As Double
    cleargrid grdcusledger: InitCRJPayment
    Screen.MousePointer = 11
    totalpayment = 0
    test = "8/31/2008"
    'Set RSSJVoucherno = gconDMIS.Execute("SELECT  distinct (voucherno),acctname FROM amis_vw_araging where (jtype = 'SJ' or jtype = 'COB') and acct_code='" & returnAcctCode(cboAccount.Text) & "'  order by voucherno asc")
    Set RSSJVoucherno = gconDMIS.Execute("SELECT  voucherno,invoicetype,invoicedate,jdate,jtype,debit FROM amis_journal_hd where (jtype = 'SJ' or jtype = 'COB') order by voucherno asc")
    Picture4.Visible = False
    If Not RSSJVoucherno.EOF And Not RSSJVoucherno.BOF Then
        RSSJVoucherno.MoveFirst
        Do While Not RSSJVoucherno.EOF
            PROGBAR.Value = 0
            PROGBAR.Max = RSSJVoucherno.RecordCount
            Picture4.Visible = True
            'Start
            'Set RsJournalHeader = gconDMIS.Execute("Select * from AMIS_vw_araging where voucherno='" & (RSSJVoucherno!VOUCHERNO) & "' and (jtype = 'SJ' or jtype = 'COB')  and acct_code='" & returnAcctCode(cboAccount.Text) & "'")
            Set RsJournalHeader = gconDMIS.Execute("Select * from AMIS_journal_hd where voucherno='" & (RSSJVoucherno!VOUCHERNO) & "' and (jtype = 'SJ' or jtype = 'COB')")
            If Not RsJournalHeader.EOF And Not RsJournalHeader.BOF Then
                RsJournalHeader.MoveFirst
                Do While Not RsJournalHeader.EOF
                    cnt = cnt + 1
                    'grdcusledger.AddItem (RsJournalHeader!Jdate) & Chr(9) & (RsJournalHeader!VOUCHERNO) & Chr(9) & _
                     '(RsJournalHeader!jtype) & Chr(9) & (RsJournalHeader!InvoiceDate) & Chr(9) & _
                     '(RsJournalHeader!InvoiceType) & Chr(9) & (RsJournalHeader!InvoiceNo) & Chr(9) & _
                     'ToDoubleNumber(RsJournalHeader!DEBIT)
                    BALANCE = NumericVal(RsJournalHeader!DEBIT)
                    totalpayment = 0
                    Set RsCrjDetail = gconDMIS.Execute("Select * from amis_crjdetail_total where invoiceno='" & (RsJournalHeader!INVOICENO) & "' and invoicetype='" & (RsJournalHeader!InvoiceType) & "' and j_class='" & returnAcctCode(cboAccount.Text) & "' and jdate <= ('" & test & "')")
                    Do While Not RsCrjDetail.EOF
                        'grdcusledger.AddItem ReturnCRJHeader(RsCrjDetail!VOUCHERNO) & Chr(9) & (RsCrjDetail!VOUCHERNO) & Chr(9) & _
                         '(RsCrjDetail!CR_type) & Chr(9) & (RsCrjDetail!InvoiceDate) & Chr(9) & _
                         '(RsCrjDetail!InvoiceType) & Chr(9) & (RsCrjDetail!InvoiceNo) & Chr(9) & _
                         'ToDoubleNumber(NumericVal(RsCrjDetail!INVOICEAMOUNT))

                        totalpayment = totalpayment + NumericVal(RsCrjDetail!invoiceamount)
                        RsCrjDetail.MoveNext
                    Loop


                    totalbalance = NumericVal(BALANCE) - NumericVal(totalpayment)

                    gconDMIS.Execute ("INSERT INTO AMIS_AR_AGING(voucherno,jtype,customercode,invoicetype,invoiceno,invoiceamt,jdate,balance) VALUES('" & (RSSJVoucherno!VOUCHERNO) & _
                                      "','SJ','" & (RsJournalHeader!CustomerCode) & "','" & (RsJournalHeader!InvoiceType) & "','" & (RsJournalHeader!INVOICENO) & _
                                      "','" & (RsJournalHeader!DEBIT) & "','" & (RsJournalHeader!JDate) & _
                                      "','" & totalbalance & "')")
                    RsJournalHeader.MoveNext
                Loop

            End If
            RSSJVoucherno.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
    MsgBox "tapos"
    Set RSSJVoucherno = Nothing
End Sub
Function CheckCRJCustomer(CRJVoucherno As String, XCustomerCode As String) As Boolean
    Dim RS                                        As New ADODB.Recordset
    Set RS = gconDMIS.Execute("Select Customercode from AMIS_journal_hd where voucherno = '" & CRJVoucherno & "' and jtype = 'CRJ'")
    If Not (RS.BOF And RS.EOF) Then
        If RS!CustomerCode <> XCustomerCode Then
            CheckCRJCustomer = True                        ' customer code is not equal from SJ to CRJ
        Else
            CheckCRJCustomer = False                       'parehas po
        End If
    End If
    Set RS = Nothing
End Function
Sub ar()
    Dim rsHeader                                  As New ADODB.Recordset
    Dim rsdetail                                  As New ADODB.Recordset
    Dim BALANCE                                   As Double
    Dim totalpayment                              As Double
    Dim CRJVoucher                                As String
    Dim Reference                                 As String
    Dim SystemRemarks                             As String
    Dim CRJInvoiceno                              As String
    Dim CRJInvoicetype                            As String
    Dim AMOUNT2PAY                                As Double
    'Set rsHeader = gconDMIS.Execute("Select Voucherno,invoicetype,invoiceno,jtype,invoicedate,invoiceamt,debit,customercode as SJ_CustomerCode from AMIS_journal_hd where (jtype = 'SJ' or jtype = 'COB') order by voucherno asc")
    Set rsHeader = gconDMIS.Execute("SELECT dbo.AMIS_Journal_HD.VoucherNo,dbo.AMIS_Journal_HD.status, dbo.AMIS_Journal_HD.JType, dbo.AMIS_Journal_HD.CustomerCode as SJ_CustomerCode,dbo.AMIS_Journal_HD.InvoiceType, dbo.AMIS_Journal_HD.InvoiceNo, dbo.AMIS_Journal_HD.InvoiceDate, dbo.AMIS_Journal_HD.InvoiceAmt,dbo.AMIS_Journal_HD.AmountToPay, dbo.AMIS_Journal_HD.AmountPaid, dbo.AMIS_Journal_Det.Acct_Code as acct_code, dbo.AMIS_Journal_Det.Acct_Name,dbo.AMIS_Journal_Det.Debit as Detdebit FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType WHERE (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') AND (dbo.AMIS_Journal_HD.JType = N'SJ' OR dbo.AMIS_Journal_HD.JType = N'COB') OR (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03') and dbo.AMIS_Journal_HD.status = 'P'  ORDER BY dbo.AMIS_Journal_HD.VoucherNo")
    rsHeader.MoveFirst
    PROGBAR.Value = 0
    PROGBAR.Max = rsHeader.RecordCount
    Picture4.Visible = True
    gconDMIS.Execute ("truncate table amis_ar")
    Do While Not rsHeader.EOF
        Reference = (rsHeader!jtype) + "-" + (rsHeader!VOUCHERNO)
        If (rsHeader!jtype) = "SJ" Then
            AMOUNT2PAY = N2Str2Zero(rsHeader!detdebit)
        Else
            AMOUNT2PAY = (rsHeader!InvoiceAmt)
        End If
        Set rsdetail = gconDMIS.Execute("Select invoiceno,invoicetype,invoiceamount,voucherno,customercode as CRJ_customercode,acct_code as CRJAcct_code from AMIS_crjdetail_total where invoicetype='" & Null2String(rsHeader!InvoiceType) & "' and invoiceno='" & rsHeader!INVOICENO & "' and status = 'P'")
        'Set Rsdetail = gconDMIS.Execute("SELECT dbo.AMIS_CRJ_Detail.INVOICEAMOUNT, dbo.AMIS_CRJ_Detail.INVOICENO, dbo.AMIS_CRJ_Detail.INVOICETYPE, dbo.AMIS_CRJ_Detail.INVOICEDATE,dbo.AMIS_Journal_HD.JDate, dbo.AMIS_Journal_HD.JType, dbo.AMIS_CRJ_Detail.VoucherNo, dbo.AMIS_Journal_HD.InvoiceNo AS HDInvoice,dbo.AMIS_CRJ_Detail.J_Class, dbo.AMIS_CRJ_Detail.SJ_voucherno, dbo.AMIS_Journal_HD.CustomerCode as CRJ_customercode, dbo.AMIS_Journal_HD.Status,dbo.AMIS_CRJ_Detail.CR_TYPE FROM dbo.AMIS_CRJ_Detail INNER JOIN dbo.AMIS_Journal_HD ON dbo.AMIS_CRJ_Detail.CR_TYPE = dbo.AMIS_Journal_HD.JType AND dbo.AMIS_CRJ_Detail.VoucherNo = dbo.AMIS_Journal_HD.VoucherNo")

        If Not rsdetail.EOF And Not rsdetail.BOF Then
            rsdetail.MoveFirst
            Do While Not rsdetail.EOF
                If (Null2String(rsHeader!SJ_CustomerCode) = Null2String(rsdetail!CRJ_customercode) And Null2String(rsHeader!Acct_code) = Null2String(rsdetail!CRJacct_code)) Then
                    CRJVoucher = Null2String(rsdetail!VOUCHERNO)
                    CRJInvoiceno = Null2String(rsdetail!InvoiceType)
                    CRJInvoicetype = Null2String(rsdetail!INVOICENO)
                    totalpayment = totalpayment + NumericVal(rsdetail!invoiceamount)
                    SystemRemarks = "NULL"
                Else
                    CRJVoucher = Null2String(rsdetail!VOUCHERNO)
                    SystemRemarks = "Wrong customer code "
                End If
                rsdetail.MoveNext
            Loop
        Else
            CRJVoucher = "NULL"
            CRJInvoiceno = "NULL"
            CRJInvoicetype = "NULL"
        End If
        BALANCE = NumericVal(AMOUNT2PAY) - NumericVal(totalpayment)
        gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                          "','" & CRJVoucher & "','" & CRJInvoicetype & "','" & CRJInvoiceno & _
                          "','" & rsHeader!SJ_CustomerCode & "','" & NumericVal(AMOUNT2PAY) & "','" & NumericVal(totalpayment) & _
                          "','" & NumericVal(BALANCE) & "','" & Null2String(rsHeader!Acct_code) & "','" & SystemRemarks & "')")
        BALANCE = 0
        totalpayment = 0
        DoEvents
        PROGBAR.Value = PROGBAR.Value + 1
        Label22(0).Caption = Round((PROGBAR.Value / PROGBAR.Max * 100), 0) & "%"
        rsHeader.MoveNext
    Loop
    MsgBox "tapos na "
    Picture4.Visible = False
    Set rsHeader = Nothing
End Sub


