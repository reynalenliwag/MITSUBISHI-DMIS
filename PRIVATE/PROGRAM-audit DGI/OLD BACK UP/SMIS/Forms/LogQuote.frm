VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMISLogQuote 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   10155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogQuote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   10155
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   435
      Left            =   7050
      TabIndex        =   85
      Top             =   6750
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   435
      Left            =   9090
      TabIndex        =   84
      Top             =   6750
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   8100
      TabIndex        =   83
      Top             =   6750
      Width           =   945
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7725
      Top             =   5550
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   3  'Vertical Line
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   0
      ScaleHeight     =   2490
      ScaleWidth      =   10155
      TabIndex        =   1
      Top             =   0
      Width           =   10155
      Begin VB.TextBox txtQuotationCode 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         Top             =   315
         Width           =   3780
      End
      Begin VB.CommandButton cmdAddVehicles 
         Caption         =   "Add Vehicles"
         Height          =   330
         Left            =   3975
         TabIndex        =   18
         Top             =   315
         Width           =   1230
      End
      Begin VB.TextBox txtNotes 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         Left            =   150
         MaxLength       =   220
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   900
         Width           =   5520
      End
      Begin VB.PictureBox picNames 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   5760
         ScaleHeight     =   2385
         ScaleWidth      =   4320
         TabIndex        =   3
         Top             =   0
         Width           =   4350
         Begin XtremeShortcutBar.ShortcutCaption capCustomerDetails 
            Height          =   285
            Left            =   30
            TabIndex        =   78
            Top             =   0
            Width           =   12735
            _Version        =   655364
            _ExtentX        =   22463
            _ExtentY        =   503
            _StockProps     =   14
            Caption         =   "::Customer Information::"
            ForeColor       =   8421504
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
            ForeColor       =   8421504
         End
         Begin VB.Label lblEmail 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1500
            TabIndex        =   13
            Top             =   2085
            Width           =   2805
         End
         Begin VB.Label lblContactNo 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1500
            TabIndex        =   12
            Top             =   1800
            Width           =   2805
         End
         Begin VB.Label lblAddress 
            BackColor       =   &H00C0C0C0&
            Height          =   690
            Left            =   15
            TabIndex        =   11
            Top             =   1090
            Width           =   4290
         End
         Begin VB.Label lblContacPerson 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1500
            TabIndex        =   10
            Top             =   570
            Width           =   2805
         End
         Begin VB.Label lblProspectName 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1500
            TabIndex        =   9
            Top             =   285
            Width           =   2805
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   4
            Left            =   15
            TabIndex        =   8
            Top             =   2085
            Width           =   1470
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "Contact No:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   3
            Left            =   15
            TabIndex        =   7
            Top             =   1800
            Width           =   1470
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   225
            Index           =   2
            Left            =   15
            TabIndex        =   6
            Top             =   855
            Width           =   4290
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "Contact Person"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   1
            Left            =   15
            TabIndex        =   5
            Top             =   570
            Width           =   1470
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "Prospect Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   0
            Left            =   15
            TabIndex        =   4
            Top             =   285
            Width           =   1470
         End
      End
      Begin VB.Label zlblC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00750A04&
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   675
         Width           =   750
      End
      Begin VB.Label zlblC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quotation Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00750A04&
         Height          =   210
         Index           =   9
         Left            =   150
         TabIndex        =   2
         Top             =   75
         Width           =   1275
      End
   End
   Begin VB.PictureBox picDetails 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   0
      ScaleHeight     =   2505
      ScaleWidth      =   10155
      TabIndex        =   14
      Top             =   2490
      Width           =   10155
      Begin XtremeReportControl.ReportControl lvQuotationVehicles 
         Height          =   2355
         Left            =   3540
         TabIndex        =   15
         Top             =   90
         Width           =   6540
         _Version        =   655364
         _ExtentX        =   11536
         _ExtentY        =   4154
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         AllowColumnResize=   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picVehiclesQDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   2355
         Left            =   150
         ScaleHeight     =   2325
         ScaleWidth      =   3330
         TabIndex        =   61
         Top             =   90
         Width           =   3360
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "Net Monthly Amortization"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   600
            Index           =   10
            Left            =   0
            TabIndex        =   74
            Top             =   1710
            Width           =   1470
         End
         Begin VB.Label lblVQNetMonthly 
            BackColor       =   &H00C0C0C0&
            Height          =   600
            Left            =   1485
            TabIndex        =   73
            Top             =   1710
            Width           =   2805
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "Terms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   8
            Left            =   0
            TabIndex        =   72
            Top             =   855
            Width           =   1470
         End
         Begin VB.Label lblVQTerms 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1485
            TabIndex        =   71
            Top             =   855
            Width           =   2805
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "DownPayment"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   7
            Left            =   0
            TabIndex        =   70
            Top             =   576
            Width           =   1470
         End
         Begin VB.Label lblVQDownPayment 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1485
            TabIndex        =   69
            Top             =   576
            Width           =   2805
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "CODE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   9
            Left            =   0
            TabIndex        =   68
            Top             =   285
            Width           =   1470
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "AOR"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   6
            Left            =   0
            TabIndex        =   67
            Top             =   1140
            Width           =   1470
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "Bal To Financed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   5
            Left            =   0
            TabIndex        =   66
            Top             =   1425
            Width           =   1470
         End
         Begin VB.Label lblVQCode 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1485
            TabIndex        =   65
            Top             =   285
            Width           =   2805
         End
         Begin VB.Label lblVQAOR 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1485
            TabIndex        =   64
            Top             =   1140
            Width           =   2805
         End
         Begin VB.Label lblVQBalToFinanced 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1485
            TabIndex        =   63
            Top             =   1425
            Width           =   2805
         End
         Begin VB.Label lblProfileType 
            Appearance      =   0  'Flat
            BackColor       =   &H00750A04&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "::Vehicle Quotation Detail::"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   1
            Left            =   0
            TabIndex        =   62
            Top             =   0
            Width           =   5205
         End
      End
   End
   Begin VB.PictureBox picQuoteFooter 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1665
      Left            =   0
      ScaleHeight     =   1665
      ScaleWidth      =   10155
      TabIndex        =   80
      Top             =   4995
      Width           =   10155
      Begin VB.TextBox txtQfoot 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   75
         MaxLength       =   220
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   81
         Top             =   225
         Width           =   9945
      End
      Begin VB.Label zlblC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Footer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00750A04&
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   82
         Top             =   -30
         Width           =   615
      End
   End
   Begin VB.PictureBox picVehiclesDetails 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   5595
      Left            =   1500
      ScaleHeight     =   5565
      ScaleWidth      =   6990
      TabIndex        =   19
      Top             =   750
      Visible         =   0   'False
      Width           =   7020
      Begin VB.TextBox txtVNetRate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3660
         TabIndex        =   79
         Text            =   "0"
         Top             =   1155
         Width           =   3165
      End
      Begin VB.TextBox txtvNetMonthlyMort 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3645
         TabIndex        =   38
         Text            =   "0"
         Top             =   4680
         Width           =   2970
      End
      Begin VB.TextBox txtVBalToFin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3645
         TabIndex        =   37
         Text            =   "0"
         Top             =   4080
         Width           =   3000
      End
      Begin VB.TextBox txtVAOR 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3645
         TabIndex        =   36
         Text            =   "0"
         Top             =   3480
         Width           =   3015
      End
      Begin VB.TextBox txtVdownpayment 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3645
         TabIndex        =   33
         Text            =   "0"
         Top             =   2340
         Width           =   3135
      End
      Begin VB.TextBox txtVTotalAmount 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3645
         TabIndex        =   32
         Top             =   1755
         Width           =   3150
      End
      Begin VB.TextBox txtVQty 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   31
         Text            =   "1"
         Top             =   600
         Width           =   510
      End
      Begin VB.CommandButton cmdCancelDetailVehicles 
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
         Left            =   6630
         TabIndex        =   60
         Top             =   30
         Width           =   285
      End
      Begin VB.PictureBox picVInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4500
         Left            =   60
         ScaleHeight     =   4470
         ScaleWidth      =   3510
         TabIndex        =   40
         Top             =   360
         Width           =   3540
         Begin VB.Label lblvModel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1485
            TabIndex        =   76
            Top             =   2175
            Width           =   1995
         End
         Begin VB.Label lblvDescript 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1365
            Left            =   0
            TabIndex        =   75
            Top             =   795
            Width           =   3480
         End
         Begin VB.Label lblCapDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00ECCABD&
            Caption         =   "Source"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   9
            Left            =   0
            TabIndex        =   59
            Top             =   4170
            Width           =   1470
         End
         Begin VB.Label lblvSource 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1485
            TabIndex        =   58
            Top             =   4170
            Width           =   1995
         End
         Begin VB.Label lblvClass 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1485
            TabIndex        =   57
            Top             =   3885
            Width           =   1995
         End
         Begin VB.Label lblCapDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00ECCABD&
            Caption         =   "Class"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   8
            Left            =   0
            TabIndex        =   56
            Top             =   3885
            Width           =   1470
         End
         Begin VB.Label lblCapDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00ECCABD&
            Caption         =   "Vin"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   0
            TabIndex        =   55
            Top             =   3600
            Width           =   1470
         End
         Begin VB.Label lblvVin 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1485
            TabIndex        =   54
            Top             =   3600
            Width           =   1995
         End
         Begin VB.Label lblCapDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00ECCABD&
            Caption         =   "Serial No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   0
            TabIndex        =   53
            Top             =   3315
            Width           =   1470
         End
         Begin VB.Label lblvSerialNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1485
            TabIndex        =   52
            Top             =   3315
            Width           =   1995
         End
         Begin VB.Label lblCapDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00ECCABD&
            Caption         =   "Color"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   0
            TabIndex        =   51
            Top             =   3030
            Width           =   1470
         End
         Begin VB.Label lblvColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1485
            TabIndex        =   50
            Top             =   3030
            Width           =   1995
         End
         Begin VB.Label lblCapDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00ECCABD&
            Caption         =   "Year"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   0
            TabIndex        =   49
            Top             =   2745
            Width           =   1470
         End
         Begin VB.Label lblvYear 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1485
            TabIndex        =   48
            Top             =   2745
            Width           =   1995
         End
         Begin VB.Label lblCapDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00ECCABD&
            Caption         =   " Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   0
            TabIndex        =   47
            Top             =   270
            Width           =   1470
         End
         Begin VB.Label lblCapDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00ECCABD&
            Caption         =   " Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   46
            Top             =   555
            Width           =   3480
         End
         Begin VB.Label lblCapDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00ECCABD&
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   0
            TabIndex        =   45
            Top             =   2175
            Width           =   1470
         End
         Begin VB.Label lblCapDetails 
            Appearance      =   0  'Flat
            BackColor       =   &H00ECCABD&
            Caption         =   "Make"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   0
            TabIndex        =   44
            Top             =   2460
            Width           =   1470
         End
         Begin VB.Label lblvCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1485
            TabIndex        =   43
            Top             =   270
            Width           =   1995
         End
         Begin VB.Label lblvMake 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1485
            TabIndex        =   42
            Top             =   2460
            Width           =   1995
         End
         Begin VB.Label zlblC 
            Appearance      =   0  'Flat
            BackColor       =   &H00750A04&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "   ::Vehicles Detail ::"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   15
            Left            =   -30
            TabIndex        =   41
            Top             =   -30
            Width           =   3555
         End
      End
      Begin VB.ComboBox cboVehicles 
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
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   3645
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   585
         Width           =   2625
      End
      Begin VB.ComboBox cboVTerm 
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
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   3645
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   2925
         Width           =   3210
      End
      Begin VB.CommandButton cmdOkVehicles 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5970
         MaskColor       =   &H00000040&
         Picture         =   "LogQuote.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5085
         Width           =   420
      End
      Begin VB.CommandButton cmdCancelDetailVehicles 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   6465
         MaskColor       =   &H00000040&
         Picture         =   "LogQuote.frx":06F6
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5085
         Width           =   420
      End
      Begin XtremeShortcutBar.ShortcutCaption caption 
         Height          =   285
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   7020
         _Version        =   655364
         _ExtentX        =   12382
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "::: Add Vehicles :::"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         Alignment       =   1
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   3645
         TabIndex        =   39
         Top             =   3285
         Width           =   375
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   7
         Left            =   3645
         TabIndex        =   29
         Top             =   1530
         Width           =   1125
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   6
         Left            =   3645
         TabIndex        =   28
         Top             =   945
         Width           =   720
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bal. to be financed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   10
         Left            =   3645
         TabIndex        =   27
         Top             =   3825
         Width           =   1560
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Mo. Amort."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   8
         Left            =   3645
         TabIndex        =   26
         Top             =   4410
         Width           =   1245
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Terms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   3645
         TabIndex        =   25
         Top             =   2700
         Width           =   555
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicles Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   3645
         TabIndex        =   24
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   5
         Left            =   6345
         TabIndex        =   23
         Top             =   360
         Width           =   285
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Down Payment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   4
         Left            =   3645
         TabIndex        =   22
         Top             =   2115
         Width           =   1275
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption capMain 
      Height          =   315
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Width           =   12795
      _Version        =   655364
      _ExtentX        =   22569
      _ExtentY        =   556
      _StockProps     =   14
      Caption         =   "Quotations"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
   End
   Begin VB.Menu mnuLook 
      Caption         =   "Look"
      Visible         =   0   'False
      Begin VB.Menu mnuUnit 
         Caption         =   "Unit In Stock"
      End
   End
End
Attribute VB_Name = "frmSMISLogQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QuotationID                             As Long
Dim ProfileID                               As Long
Dim ProfileType                             As String
Dim ProspectID                              As Long
Dim AcctName                                As String
Dim ListEdit                                As Boolean
Dim RsVehicles                              As Recordset





Private Sub cboVehicles_Click()
    If cboVehicles.ListIndex = -1 Then: Exit Sub
    Dim temprs                              As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT * From CRIS_MRRINV WHERE ID=" & cboVehicles.ItemData(cboVehicles.ListIndex))
    If Not (temprs.EOF Or temprs.BOF) Then
        txtVNetRate.Text = N2Str2IntZero(temprs!taggedprice)
        lblvCode = Null2String(temprs!code)
        lblvDescript = Null2String(temprs!descript)
        lblvModel = Null2String(temprs!Model)
        lblvMake = Null2String(temprs!Make)
        lblvYear = Null2String(temprs!yearmodel)
        lblvColor = Null2String(temprs!Color)
        lblvSerialNo = Null2String(temprs!serialno)
        lblvVin = Null2String(temprs!vinnumber)
        lblvClass = Null2String(temprs!Class)
        lblvSource = Null2String(temprs!Source)

    End If
    Set temprs = Nothing
End Sub

'

Private Sub cboVTerm_Click()
    Dim NoOfMonths                          As Integer
    Dim AOR                                 As Double
    AOR = cboVTerm.ItemData(cboVTerm.ListIndex)
    NoOfMonths = Left(cboVTerm.Text, 2)
    If NoOfMonths = 0 Then AOR = 0: NoOfMonths = 1
    If NoOfMonths = 12 Then AOR = 7.61
    If NoOfMonths = 18 Then AOR = 10.48
    If NoOfMonths = 24 Then AOR = 17.45
    If NoOfMonths = 36 Then AOR = 25.55
    If NoOfMonths = 48 Then AOR = 33.96
    If NoOfMonths = 60 Then AOR = 44.15
    txtVAOR.Text = NumericVal(AOR)
    txtVBalToFin.Text = NumericVal(txtVTotalAmount.Text) - NumericVal(txtVdownpayment.Text)
    txtvNetMonthlyMort.Text = ((NumericVal(txtVBalToFin.Text)) * (1 + (AOR / 100)) / NoOfMonths)
End Sub

Sub CenterPicture(picx As PictureBox)
    picx.Left = (Me.ScaleWidth - picx.Width) / 2
    picx.Top = (Me.ScaleHeight - picx.Height) / 2
End Sub


Private Sub cmdAddVehicles_Click()
    ListEdit = False
    cboVehicles.Enabled = True
    ShowHidePictureBox picVehiclesDetails.hwnd, True, Me
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelDetailVehicles_Click(Index As Integer)
    ShowHidePictureBox picVehiclesDetails.hwnd, False, Me
End Sub

Private Sub cmdOkVehicles_Click()
    If txtVNetRate.Text = 0 Then: Call ColorIt(txtVNetRate, Timer1): Exit Sub
    If txtVQty.Text = 0 Then: Call ColorIt(txtVQty, Timer1): Exit Sub



    Dim lst                                 As ReportRecord
    Dim totamount                           As Currency
    Dim i                                   As Integer
    Dim REC                                 As ReportRecord
    Dim ItemExist                           As Boolean
    ''FIND ITEM


    If ListEdit = True Then
        Set lst = lvQuotationVehicles.SelectedRows(0).Record
        prc_FillLines lst

        Set lst = Nothing
        ShowHidePictureBox picVehiclesDetails.hwnd, False, Me
        Exit Sub
    End If

    For i = 0 To lvQuotationVehicles.Records.Count - 1
        If lblvCode.caption = lvQuotationVehicles.Records(i).Item(1).Value Then
            Set lst = lvQuotationVehicles.Records(i)
            ItemExist = True
            Exit For
        End If
    Next


    ''OBJ EXISTS
    If Not lst Is Nothing Then
        If ItemExist = True Then
            If MsgBox("Item Exists In List." & vbCrLf & "Do you Want To Update?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Item Exists") = vbYes Then
                prc_FillLines lst
                Me.Refresh
            End If
        Else
            prc_FillLines lst
        End If
    Else
        ''ADD NEW LINE
        Set REC = lvQuotationVehicles.Records.Add
        REC.AddItem (lvQuotationVehicles.Records.Count)
        REC.AddItem lblvCode.caption
        REC.AddItem lblvDescript
        REC.AddItem txtVQty.Text
        REC.AddItem txtVNetRate.Text
        REC.AddItem txtVdownpayment.Text
        REC.AddItem (cboVTerm.Text)
        REC.AddItem (txtVAOR.Text)
        REC.AddItem (txtVBalToFin.Text)
        REC.AddItem (txtvNetMonthlyMort.Text)

        lvQuotationVehicles.Populate
    End If
    ''''''UPDATE AMOUNT
    ShowHidePictureBox picVehiclesDetails.hwnd, False, Me
    Set lst = Nothing
    ''''''SET DEFAULT LINES

End Sub

Private Sub cmdSave_Click()
    Dim i                                   As Integer
    Dim SQL                                 As String
    Dim temprs                              As ADODB.Recordset
    If QuotationID <= 0 Then
        SQL = "INSERT INTO CRIS_Quote_Header(ProspectID,  QuotationCode, QuotationDescription,QFoot) " _
            & " VALUES(@ProspectID,  @QuotationCode, @QuotationDescription, @QFoot)" & vbCrLf & " SELECT @@IDENTITY"
    Else

        SQL = "  Update CRIS_Quote_Header " _
            & " SET QuotationCode=@QuotationCode, QuotationDescription=@QuotationDescription, QFoot=@QFoot " _
            & " WHERE QuotationID=@QuotationID "


    End If
    SQL = Replace(SQL, "@QuotationID", QuotationID)
    SQL = Replace(SQL, "@QuotationCode", N2Str2Null(txtQuotationCode.Text))
    SQL = Replace(SQL, "@QuotationDescription", N2Str2Null(txtNotes.Text))
    SQL = Replace(SQL, "@ProspectID", ProspectID)
    SQL = Replace(SQL, "@ProfileType", N2Str2Null(ProfileType))
    SQL = Replace(SQL, "@QFoot", N2Str2Null(txtQfoot))



    Set temprs = gconDMIS.Execute(SQL)
    Set temprs = temprs.NextRecordset
    If Not temprs Is Nothing Then
        QuotationID = temprs.Collect(0)
    End If

    gconDMIS.Execute ("DELETE FROM CRIS_Quote_Details WHERE QuotationCode=" & N2Str2Null(txtQuotationCode.Text))

    For i = 0 To lvQuotationVehicles.Records.Count - 1
        With lvQuotationVehicles.Records(i)
            SQL = "INSERT INTO CRIS_Quote_Details " & _
                "   (QuotationCode, EntryCode, Price,Qty, Downpayment, Terms,AOR, BalToFin, NetMonthlyAmort) " & _
                  "VALUES (@QuotationCode, @EntryCode, @Price,@QTY,@DownPayment, @Terms, @AOR , @BalToFin , @NetMonthlyAmort ) "

            SQL = Replace(SQL, "@QuotationCode", N2Str2Null(txtQuotationCode.Text))
            SQL = Replace(SQL, "@EntryCode", N2Str2Null(.Item(1).Value))
            SQL = Replace(SQL, "@QuotationType", "'V'")
            SQL = Replace(SQL, "@AOR", .Item(7).Value)
            SQL = Replace(SQL, "@BalToFin", CCur(.Item(8).Value))
            SQL = Replace(SQL, "@NetMonthlyAmort", CCur(.Item(9).Value))
            SQL = Replace(SQL, "@QTY", .Item(3).Value)
            SQL = Replace(SQL, "@Price", CCur(.Item(4).Value))
            SQL = Replace(SQL, "@DownPayment", CCur(.Item(5).Value))
            SQL = Replace(SQL, "@Terms", .Item(7).Value)

            gconDMIS.Execute SQL
        End With
    Next
    SQL = "UPDATE     CRIS_Quote_Header Set TOtalAmount=@TotalAmount Where QuotationID=@QuotationID"
    SQL = Replace(SQL, "@QuotationID", QuotationID)
    SQL = Replace(SQL, "@TotalAmount", NumericVal(0))
    gconDMIS.Execute SQL
    MessagePop RecSave, "Record Saved", " Record Saved "
    txtQuotationCode.Enabled = False

    gconDMIS.Execute ("Update CRIS_PROSPECTS SET LogQuote=getdate() Where ProspectID=" & ProspectID)


    MainForm.ProspectStatus.ProspectID = ProspectID
End Sub

Public Sub FillListData(objx As Object, oRs As Recordset, ShowingIndex As String)

    objx.Clear
    While Not oRs.EOF
        objx.AddItem (oRs.Fields(ShowingIndex))
        objx.ItemData(objx.NewIndex) = oRs.Collect(0)
        oRs.MoveNext
    Wend
    If objx.ListCount > 0 Then
        If TypeName(objx) = "ComboBox" Then
            objx.ListIndex = 0
        End If
    End If
End Sub



Private Sub Form_Load()

    CenterPicture picVehiclesDetails


    Set RsVehicles = gconDMIS.Execute("SELECT ID,CODE, DESCRIPT as Particulars, MODEL, MAKE, YEarModel  FROM CRIS_MRRINV")

    If (RsVehicles.EOF Or RsVehicles.BOF) = True Then
        MessagePop InfoVoid, "InSuffcient Record", "There Are No Enlisted Master File For Vehicles. Please Add Few "
    End If

    FillListData cboVehicles, RsVehicles, "Particulars"

    With lvQuotationVehicles
        .Columns.Add 0, "Item", 30, False
        .Columns.Add 1, "CODE", 80, False
        .Columns.Add 2, "Description", 165, False
        .Columns.Add 3, "QTY", 35, False
        .Columns.Add 4, "Net Price", 120, False
    End With



    With lvQuotationVehicles
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GroupRowTextBold = True         ' = vbWhite
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
    End With

    With cboVTerm
        .Clear
        .AddItem "0 mos."
        .ItemData(.NewIndex) = 0
        .AddItem "12 mos."
        .ItemData(.NewIndex) = 7.61
        .AddItem "18 mos."
        .ItemData(.NewIndex) = 10.48
        .AddItem "24 mos."
        .ItemData(.NewIndex) = 17.45
        .AddItem "36 mos."
        .ItemData(.NewIndex) = 25.55
        .AddItem "48 mos."
        .ItemData(.NewIndex) = 33.96
        .AddItem "60 mos."
        .ItemData(.NewIndex) = 44.15
    End With
    If QuotationID = 0 Then
        txtQuotationCode = GenerateCode("CRIS_QUOTE_HEADER", "QuotationCode", "0000000000")
    End If


End Sub


Private Sub Form_Unload(Cancel As Integer)
    QuotationID = 0
End Sub


Sub LabelIt()
    Dim temprs                              As ADODB.Recordset

    Set temprs = gconDMIS.Execute("select * from   CRIS_PROSPECTS where PROSPECTID=" & ProspectID)
    If Not (temprs.EOF Or temprs.BOF) Then
        lblProspectName.caption = Null2String(temprs!AcctName)
        lblContacPerson.caption = Null2String(temprs!contactperson)
        lblAddress.caption = Null2String(temprs!Address)
        lblContactNo.caption = Null2String(temprs!Telephone) & "/" & Null2String(temprs!Mobile)
        lblEmail.caption = Null2String(temprs!Email)
    End If
    Set temprs = Nothing
End Sub


Friend Sub NewQuotation(xProspectID As Long)
    ProspectID = xProspectID
    ListEdit = False
    LabelIt
End Sub

Friend Sub OpenQuotation(CustID As Long, CustType As String, QCode As String)
    ProfileID = CustID
    ProfileType = CustType
    QuotationID = QCode
    LabelIt
    prc_FillHeader
    prc_FillDetails
End Sub

Private Sub prc_FillDetails()

    '
    '        With lvQuotation
    '        .Columns.Add 0, "Item", 30, False
    '        .Columns.Add 1, "CODE", 90, False
    '        .Columns.Add 2, "Type", 100, False
    '        .Columns.Add 3, "Description", 250, False
    '        .Columns.Add 4, "QTY", 35, False
    '        .Columns.Add 5, "Rate", 50, False
    '        .Columns.Add 6, "Amount", 120, False
    '        .GroupsOrder.Add .Columns(2)
    '        .Columns(2).Visible = False
    '    End With

    '    End With
    Dim temprs                              As ADODB.Recordset
    Dim SQL                                 As String
    Dim fld                                 As Field
    Dim j                                   As Long
    Dim rec1                                As XtremeReportControl.ReportRecord
    Dim rec2                                As XtremeReportControl.ReportRecord


    'Sql = "SELECT  " & _
     "EntryCode as CODE ,  " & _
     "case  QuotationType  " & _
     "WHEN 'P' then 'Parts' " & _
     "WHEN 'V' then 'Vehicles' " & _
     "WHEN 'S' then 'Services' " & _
     "WHEN 'M' then 'Materials' " & _
     "END as TYPE , " & _
     "case  QuotationType  " & _
     "WHEN 'P' then (Select TOP 1 ISNULL(STOCKDESC,'N/A') from PMIS_STOCKMAS WHERE STOCKNO=EntryCode)  " & _
     "WHEN 'V' then (SELECT TOP 1 DESCRIPT FROM ALL_Model WHERE CODE=EntryCode)  " & _
     "WHEN 'S' then (SELECT TOP 1 CJ.Desc1 as Particulars FROM CSMS_Jobs CJ WHERE JCODE=EntryCode)  " & _
     "WHEN 'M' then (Select TOP 1 STOCKDESC from PMIS_STOCKMAS WHERE STOCKNO=EntryCode)  " & _
     "END as [Description] , " & _
     "QTY,  " & _
     "Price, " & _
     "QTY * PRICE  as AMOUNT , DownPayment,Terms , AOR , BalToFin , NetMonthlyAmort " & _
     "FROM  " & _
     "CRIS_Quote_Details  " & _
     "WHERE QuotationCode=" & N2Str2Null(txtQuotationCode.Text)

    SQL = "Select Code, Type, [Description], Qty, Price, Amount From CRIS_Vw_QuotationDetails Where Type<>'Vehicles' and QuotationCode=" & N2Str2Null(txtQuotationCode.Text)
    Set temprs = gconDMIS.Execute(SQL)

    '    While Not temprs.EOF
    '        j = j + 1
    '        Set rec1 = lvQuotation.Records.Add
    '        rec1.AddItem j
    '        For Each fld In temprs.Fields
    '            rec1.AddItem (Trim(fld.Value))
    '        Next
    '        temprs.MoveNext
    '    Wend
    'Code               0
    '[Description]      1
    'Qty                2
    'Price              3
    'DownPayment        4
    'Terms              5
    'AOR                6
    'BalToFin           7
    'NetMonthlyAmort    8


    lvQuotationVehicles.Records.DeleteAll

    'QuotationCOde, EntryCode, Qty, Price, Downpayment, Terms, AOR, BalToFin, NetMonthlyAMort
    SQL = "Select Code, [Description], Qty, Price,  DownPayment, Terms, AOR, BalToFin, NetMonthlyAmort From CRIS_Vw_QuotationDetails  Where Type='Vehicles' and QuotationCode=" & N2Str2Null(txtQuotationCode.Text)
    Set temprs = gconDMIS.Execute(SQL)
    j = 0
    While Not temprs.EOF
        j = j + 1
        Set rec2 = lvQuotationVehicles.Records.Add
        rec2.AddItem j
        For Each fld In temprs.Fields
            If fld.Type = adDouble Then
                rec2.AddItem (FormatCurrency(fld.Value, 2, vbTrue, vbTrue, vbTrue))
            Else
                rec2.AddItem (Trim(fld.Value))
            End If

        Next
        temprs.MoveNext
    Wend


    lvQuotationVehicles.Populate
    Set fld = Nothing
    Set rec1 = Nothing
    Set rec2 = Nothing
    Set temprs = Nothing
    'prc_UpdateSubTotal


End Sub

Private Sub prc_FillHeader()
    If QuotationID <= 0 Then: Exit Sub


    Dim oRsx                                As ADODB.Recordset

    Set oRsx = gconDMIS.Execute("Select * from CRIS_Quote_Header Where QuotationID=" & QuotationID)

    If Not oRsx.EOF Or oRsx.BOF Then
        txtQuotationCode.Text = oRsx.Fields("QuotationCode")
        txtNotes.Text = Null2String(oRsx.Fields("QuotationDescription"))
        txtQfoot.Text = Null2String(oRsx.Fields("Qfoot"))
        txtQuotationCode.Enabled = False
        oRsx.MoveNext
    End If
    Set oRsx = Nothing
End Sub

Private Sub prc_FillLines(lst As ReportRecord)

    With lst
        .Item(4).Value = NumericVal(txtVNetRate.Text)
        .Item(5).Value = NumericVal(txtVdownpayment.Text)
        .Item(6).Value = cboVTerm.Text
        .Item(7).Value = NumericVal(txtVAOR.Text)
        .Item(8).Value = NumericVal(txtVBalToFin.Text)
        .Item(9).Value = NumericVal(txtvNetMonthlyMort.Text)

        lvQuotationVehicles.Populate

    End With


End Sub


Private Sub lvQuotationVehicles_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    With Row
        cboVehicles.ListIndex = SelectCombo(cboVehicles, .Record(2).Value)
        cboVehicles.Enabled = False
        txtVQty = .Record(3).Value
        txtVNetRate = .Record(4).Value
        txtVdownpayment = .Record(5).Value
        cboVTerm.ListIndex = SelectCombo(cboVTerm, .Record(6).Value)
        txtVAOR = .Record(7).Value
        txtVBalToFin = .Record(8).Value
        txtvNetMonthlyMort = .Record(9).Value
    End With
    ListEdit = True
    ShowHidePictureBox picVehiclesDetails.hwnd, True, Me
End Sub

Private Sub lvQuotationVehicles_SelectionChanged()
    With lvQuotationVehicles.SelectedRows.Row(0)
        lblVQCode.caption = Space(1) & .Record(1).Value
        lblVQDownPayment.caption = Space(1) & .Record(5).Value
        lblVQTerms.caption = Space(1) & .Record(6).Value
        lblVQAOR.caption = Space(1) & .Record(7).Value
        lblVQBalToFinanced.caption = Space(1) & .Record(8).Value
        lblVQNetMonthly.caption = Space(1) & .Record(9).Value
    End With
End Sub

Private Sub Timer1_Timer()
    Dim cntrl                               As Control
    For Each cntrl In Me.ControlS
        If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
            If cntrl.ForeColor = vbYellow Then
                cntrl.ForeColor = vbBlack
                cntrl.BackColor = vbWhite
            End If

        End If
    Next
    Timer1.Enabled = False
End Sub

Private Sub txtNotes_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtVNetRate_Change()
    txtVTotalAmount = NumericVal(txtVQty) * NumericVal(txtVNetRate)
End Sub

Private Sub txtVQty_Change()
    txtVTotalAmount = NumericVal(txtVQty) * NumericVal(txtVNetRate)
End Sub

