VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Trans_VehicleInvoice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Invocing"
   ClientHeight    =   9120
   ClientLeft      =   2430
   ClientTop       =   1755
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "VehicleInvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   10020
   Begin VB.CommandButton Command2 
      Caption         =   "&Charge"
      Height          =   615
      Left            =   1080
      MouseIcon       =   "VehicleInvoice.frx":08CA
      MousePointer    =   99  'Custom
      Picture         =   "VehicleInvoice.frx":0A1C
      Style           =   1  'Graphical
      TabIndex        =   289
      ToolTipText     =   "Print this Record"
      Top             =   7140
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Free"
      Height          =   615
      Left            =   270
      MouseIcon       =   "VehicleInvoice.frx":0D82
      MousePointer    =   99  'Custom
      Picture         =   "VehicleInvoice.frx":0ED4
      Style           =   1  'Graphical
      TabIndex        =   287
      ToolTipText     =   "Print this Record"
      Top             =   7140
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox Picture6 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   10020
      TabIndex        =   264
      Top             =   8775
      Width           =   10020
      Begin VB.Label labDetails 
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
         Left            =   4020
         TabIndex        =   269
         Top             =   0
         Width           =   5955
      End
      Begin VB.Label labSJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2880
         TabIndex        =   268
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SJ #:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   2010
         TabIndex        =   267
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " OR #:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         TabIndex        =   265
         Top             =   0
         Width           =   855
      End
      Begin VB.Label labORNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   870
         TabIndex        =   266
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.PictureBox picHeader 
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   0
      ScaleHeight     =   1605
      ScaleWidth      =   10080
      TabIndex        =   11
      Top             =   0
      Width           =   10080
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   9870
         Top             =   210
      End
      Begin VB.Frame Frame3 
         Height          =   660
         Left            =   60
         TabIndex        =   12
         Top             =   -30
         Width           =   9900
         Begin VB.TextBox txtSODate 
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
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   271
            Top             =   180
            Width           =   1680
         End
         Begin VB.TextBox txtVINo 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   465
            Left            =   720
            MaxLength       =   6
            TabIndex        =   14
            Text            =   "000000"
            Top             =   150
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SO Date"
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
            Index           =   15
            Left            =   7290
            TabIndex        =   272
            Top             =   240
            Width           =   690
         End
         Begin VB.Label LABALLOWREPRINT 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   8190
            TabIndex        =   17
            Top             =   150
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label labVDRNo 
            Caption         =   "0000000"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2880
            TabIndex        =   16
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label labInvoiceStatus 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   4200
            TabIndex        =   18
            Top             =   210
            Width           =   2895
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VI NO#"
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
            Left            =   105
            TabIndex        =   13
            Top             =   255
            Width           =   570
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VDR NO#"
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
            Left            =   1980
            TabIndex        =   15
            Top             =   270
            Width           =   765
         End
      End
      Begin VB.Frame fraHeader 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   60
         TabIndex        =   19
         Top             =   570
         Width           =   9915
         Begin VB.TextBox txtBankPo 
            BackColor       =   &H00FFFFFF&
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
            Height          =   345
            Left            =   5490
            MaxLength       =   20
            TabIndex        =   290
            Top             =   180
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.CommandButton Command4 
            Caption         =   "::"
            Height          =   345
            Left            =   9510
            TabIndex        =   30
            ToolTipText     =   "Edit Transaction Date For The Transaction"
            Top             =   570
            Width           =   345
         End
         Begin VB.ComboBox cboSalesOrderNo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8085
            TabIndex        =   25
            Top             =   172
            Width           =   1755
         End
         Begin VB.ComboBox cboPaymentTerm 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00400000&
            Height          =   345
            ItemData        =   "VehicleInvoice.frx":123A
            Left            =   2850
            List            =   "VehicleInvoice.frx":123C
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   180
            Width           =   2085
         End
         Begin VB.ComboBox cboSalesAE 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   1305
            TabIndex        =   27
            Top             =   570
            Width           =   3645
         End
         Begin VB.ComboBox cboPurchaseType 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00400000&
            Height          =   345
            ItemData        =   "VehicleInvoice.frx":123E
            Left            =   1320
            List            =   "VehicleInvoice.frx":1240
            TabIndex        =   21
            Text            =   "Combo1"
            Top             =   180
            Width           =   1020
         End
         Begin MSComCtl2.DTPicker dtDateInvoiced 
            Height          =   345
            Left            =   8085
            TabIndex        =   29
            Top             =   570
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            Format          =   54329345
            CurrentDate     =   39213
         End
         Begin MSComCtl2.DTPicker dtbankcom_po 
            Height          =   345
            Left            =   5490
            TabIndex        =   291
            Top             =   570
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            Format          =   54329345
            CurrentDate     =   39213
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Po No."
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
            Left            =   4950
            TabIndex        =   293
            Top             =   270
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label Label5 
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
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   4950
            TabIndex        =   292
            Top             =   660
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SO No#"
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
            Left            =   7425
            TabIndex        =   24
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Agent"
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
            Left            =   15
            TabIndex        =   26
            Top             =   630
            Width           =   1020
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Term"
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
            Left            =   2340
            TabIndex        =   22
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Type"
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
            Left            =   30
            TabIndex        =   20
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Date"
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
            Left            =   7050
            TabIndex        =   28
            Top             =   630
            Width           =   1005
         End
      End
   End
   Begin VB.PictureBox picPrintingDetails 
      BackColor       =   &H008080FF&
      Height          =   255
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   8835
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8895
      Begin VB.Frame fraPrintingDetails 
         BorderStyle     =   0  'None
         Caption         =   "Signatories"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   4440
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   9420
         Begin VB.TextBox txtPreparedBy 
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
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   1920
            TabIndex        =   3
            Top             =   195
            Width           =   2775
         End
         Begin VB.TextBox txtCheckedBy 
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
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   1920
            TabIndex        =   7
            Top             =   630
            Width           =   2775
         End
         Begin VB.TextBox txtSalesApproved 
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
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   6480
            TabIndex        =   9
            Top             =   675
            Width           =   2775
         End
         Begin VB.TextBox txtGeneralManager 
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
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   6480
            TabIndex        =   5
            Top             =   210
            Width           =   2775
         End
         Begin VB.Label Label48 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Notes:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   2535
            Left            =   270
            TabIndex        =   10
            Top             =   1500
            Width           =   8970
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Prepared By"
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
            Left            =   210
            TabIndex        =   2
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Checked By"
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
            Left            =   210
            TabIndex        =   6
            Top             =   675
            Width           =   1005
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Approved"
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
            Left            =   4770
            TabIndex        =   8
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "General Manager"
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
            Left            =   4770
            TabIndex        =   4
            Top             =   285
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox picAdds 
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
      Height          =   855
      Left            =   -30
      ScaleHeight     =   855
      ScaleWidth      =   11745
      TabIndex        =   248
      Top             =   7890
      Width           =   11745
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
         Left            =   60
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "VehicleInvoice.frx":1242
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":1394
         Style           =   1  'Graphical
         TabIndex        =   274
         ToolTipText     =   "Post this Transaction"
         Top             =   -30
         Visible         =   0   'False
         Width           =   315
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
         Left            =   60
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "VehicleInvoice.frx":16B9
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":180B
         Style           =   1  'Graphical
         TabIndex        =   273
         ToolTipText     =   "Unpost this Transaction"
         Top             =   -30
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         CausesValidation=   0   'False
         Height          =   795
         Left            =   9030
         MouseIcon       =   "VehicleInvoice.frx":1B50
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":1CA2
         Style           =   1  'Graphical
         TabIndex        =   260
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   8340
         MouseIcon       =   "VehicleInvoice.frx":2008
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":215A
         Style           =   1  'Graphical
         TabIndex        =   259
         ToolTipText     =   "Print this Record"
         Top             =   30
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
         Left            =   7650
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "VehicleInvoice.frx":24C0
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":2612
         Style           =   1  'Graphical
         TabIndex        =   258
         ToolTipText     =   "Cancel this Transaction"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdUnReleased 
         Caption         =   "Un Release"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   6960
         MouseIcon       =   "VehicleInvoice.frx":294C
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":2A9E
         Style           =   1  'Graphical
         TabIndex        =   255
         ToolTipText     =   "Unrelease Vehicle"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdRelease 
         Caption         =   "Release"
         Height          =   795
         Left            =   6270
         MouseIcon       =   "VehicleInvoice.frx":2E93
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":2FE5
         Style           =   1  'Graphical
         TabIndex        =   257
         ToolTipText     =   "Release Vehicle"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   5580
         MouseIcon       =   "VehicleInvoice.frx":5787
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":58D9
         Style           =   1  'Graphical
         TabIndex        =   256
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   4890
         MouseIcon       =   "VehicleInvoice.frx":5C35
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":5D87
         Style           =   1  'Graphical
         TabIndex        =   254
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   795
         Left            =   4200
         MouseIcon       =   "VehicleInvoice.frx":609A
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":61EC
         Style           =   1  'Graphical
         TabIndex        =   253
         ToolTipText     =   "Move to Last Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Height          =   795
         Left            =   3510
         MouseIcon       =   "VehicleInvoice.frx":653C
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":668E
         Style           =   1  'Graphical
         TabIndex        =   252
         ToolTipText     =   "Move to First Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   2820
         MouseIcon       =   "VehicleInvoice.frx":69EC
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":6B3E
         Style           =   1  'Graphical
         TabIndex        =   251
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   2130
         MouseIcon       =   "VehicleInvoice.frx":6E38
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":6F8A
         Style           =   1  'Graphical
         TabIndex        =   250
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Pre&v"
         Height          =   795
         Left            =   1440
         MouseIcon       =   "VehicleInvoice.frx":72E2
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":7434
         Style           =   1  'Graphical
         TabIndex        =   249
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   795
         Left            =   750
         MouseIcon       =   "VehicleInvoice.frx":7793
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":78E5
         Style           =   1  'Graphical
         TabIndex        =   275
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
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
      Left            =   8370
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   261
      Top             =   7920
      Visible         =   0   'False
      Width           =   1800
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
         Height          =   795
         Left            =   870
         MouseIcon       =   "VehicleInvoice.frx":7E60
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":7FB2
         Style           =   1  'Graphical
         TabIndex        =   262
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   180
         MouseIcon       =   "VehicleInvoice.frx":82F0
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":8442
         Style           =   1  'Graphical
         TabIndex        =   263
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   705
      End
   End
   Begin XtremeSuiteControls.TabControl SSTabVDetails 
      Height          =   5880
      Left            =   60
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1980
      Width           =   9900
      _Version        =   655364
      _ExtentX        =   17462
      _ExtentY        =   10372
      _StockProps     =   64
      Appearance      =   1
      Color           =   4
      PaintManager.Layout=   1
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   90
      ItemCount       =   4
      Item(0).Caption =   "Customers Information"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "picTinInfo"
      Item(0).Control(1)=   "picCustomerInformation"
      Item(1).Caption =   "Vehicles Detail"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "picVehiclesDetail"
      Item(2).Caption =   "Terms"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "picTerms"
      Item(3).Caption =   "Others"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "picViewAccessories"
      Begin VB.PictureBox picCustomerInformation 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   4245
         Left            =   210
         ScaleHeight     =   4245
         ScaleWidth      =   6240
         TabIndex        =   86
         Top             =   720
         Width           =   6240
         Begin VB.TextBox txtContactPerson 
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
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   1350
            TabIndex        =   92
            Top             =   600
            Width           =   4770
         End
         Begin VB.TextBox txtTelephoneHome 
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
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   3840
            TabIndex        =   106
            Top             =   3375
            Width           =   2310
         End
         Begin VB.TextBox txtTelephoneOffice 
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
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   1365
            TabIndex        =   105
            Top             =   3375
            Width           =   2460
         End
         Begin VB.TextBox txtHomeAdd 
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
            ForeColor       =   &H00701E2A&
            Height          =   750
            Left            =   1365
            MaxLength       =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   101
            Top             =   1770
            Width           =   4755
         End
         Begin VB.TextBox txtCustName 
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
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   2625
            TabIndex        =   91
            Top             =   210
            Width           =   3495
         End
         Begin VB.TextBox txtCusCode 
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
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   1365
            Locked          =   -1  'True
            TabIndex        =   90
            Tag             =   "@R"
            ToolTipText     =   "Customer Account Code"
            Top             =   210
            Width           =   1275
         End
         Begin VB.TextBox txtDateBirth 
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
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   1365
            TabIndex        =   97
            Top             =   1380
            Width           =   2250
         End
         Begin VB.TextBox txtPosition 
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
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   4410
            TabIndex        =   99
            Top             =   1395
            Width           =   1710
         End
         Begin VB.TextBox txtSpouse 
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
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   1365
            TabIndex        =   95
            Top             =   990
            Width           =   4770
         End
         Begin VB.TextBox txtOfficeAdd 
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
            ForeColor       =   &H00701E2A&
            Height          =   750
            Left            =   1365
            MaxLength       =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   103
            Top             =   2580
            Width           =   4755
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000D&
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
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   93
            Top             =   630
            Width           =   1440
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Tel.No. Office/ Home"
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
            Height          =   480
            Index           =   6
            Left            =   180
            TabIndex        =   104
            Top             =   3330
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Spouse"
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
            Index           =   2
            Left            =   675
            TabIndex        =   94
            Top             =   1050
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Birth"
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
            Index           =   3
            Left            =   270
            TabIndex        =   96
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
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
            Index           =   0
            Left            =   480
            TabIndex        =   89
            Top             =   270
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Home Address"
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
            Left            =   60
            TabIndex        =   100
            Top             =   1965
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Office Address"
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
            Left            =   45
            TabIndex        =   102
            Top             =   2760
            Width           =   1275
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Left            =   2610
            TabIndex        =   88
            Top             =   -15
            Width           =   1200
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "AC Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Left            =   1365
            TabIndex        =   87
            Top             =   -30
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
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
            Left            =   3660
            TabIndex        =   98
            Top             =   1455
            Width           =   690
         End
      End
      Begin VB.PictureBox picViewAccessories 
         BorderStyle     =   0  'None
         Height          =   5145
         Left            =   -69925
         ScaleHeight     =   5145
         ScaleWidth      =   9615
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   9615
         Begin VB.PictureBox fraAccessories 
            BorderStyle     =   0  'None
            Height          =   4995
            Left            =   0
            ScaleHeight     =   4995
            ScaleWidth      =   9540
            TabIndex        =   35
            Top             =   0
            Width           =   9540
            Begin Crystal.CrystalReport rptFree 
               Left            =   1260
               Top             =   4560
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   348160
               PrintFileLinesPerPage=   60
            End
            Begin VB.TextBox infoAdditionalInfo 
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
               Height          =   1005
               Left            =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   37
               Text            =   "VehicleInvoice.frx":8792
               Top             =   240
               Width           =   9390
            End
            Begin VB.CommandButton cmdAddAcc 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Freebies"
               Height          =   315
               Left            =   135
               MaskColor       =   &H00400000&
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   1290
               Width           =   2565
            End
            Begin MSComctlLib.ListView lvAccesories 
               Height          =   2865
               Left            =   120
               TabIndex        =   40
               Top             =   1620
               Width           =   9390
               _ExtentX        =   16563
               _ExtentY        =   5054
               View            =   3
               LabelEdit       =   1
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
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
            Begin MSMask.MaskEdBox txtTotalAccesories 
               Height          =   435
               Left            =   6675
               TabIndex        =   42
               Top             =   4530
               Width           =   2790
               _ExtentX        =   4921
               _ExtentY        =   767
               _Version        =   393216
               BackColor       =   -2147483633
               ForeColor       =   7347754
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label3 
               Caption         =   "Double Click To Edit Detail"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   4950
               TabIndex        =   39
               Top             =   1350
               Width           =   4545
            End
            Begin VB.Label Label79 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Notes:"
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
               Left            =   180
               TabIndex        =   36
               Top             =   0
               Width           =   540
            End
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
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
               Left            =   5460
               TabIndex        =   41
               Top             =   4590
               Width           =   1125
            End
         End
      End
      Begin VB.PictureBox picTinInfo 
         BorderStyle     =   0  'None
         Height          =   4635
         Left            =   6360
         ScaleHeight     =   4635
         ScaleWidth      =   3315
         TabIndex        =   107
         Top             =   660
         Width           =   3315
         Begin VB.ComboBox cboAccountType 
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
            ItemData        =   "VehicleInvoice.frx":8798
            Left            =   930
            List            =   "VehicleInvoice.frx":87A2
            TabIndex        =   117
            Top             =   1590
            Width           =   2325
         End
         Begin VB.TextBox txtDeliveryInstruction 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            MaxLength       =   30
            TabIndex        =   121
            Top             =   3780
            Width           =   3165
         End
         Begin VB.TextBox txtDeliveryAddress 
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
            ForeColor       =   &H00701E2A&
            Height          =   1125
            Left            =   90
            MaxLength       =   120
            MultiLine       =   -1  'True
            TabIndex        =   119
            Top             =   2400
            Width           =   3180
         End
         Begin VB.TextBox txtTIN 
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
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   930
            TabIndex        =   109
            Top             =   0
            Width           =   2310
         End
         Begin VB.TextBox txtIssuedon 
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
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   930
            TabIndex        =   113
            Top             =   795
            Width           =   2310
         End
         Begin VB.TextBox txtIssuedat 
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
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   930
            TabIndex        =   115
            Top             =   1185
            Width           =   2310
         End
         Begin VB.TextBox txtCTCNo 
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
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   930
            TabIndex        =   111
            Top             =   420
            Width           =   2310
         End
         Begin VB.CommandButton cmdEditCustInfo 
            Caption         =   "Edit Customer Information"
            Height          =   360
            Left            =   90
            MouseIcon       =   "VehicleInvoice.frx":87B5
            MousePointer    =   99  'Custom
            TabIndex        =   122
            TabStop         =   0   'False
            ToolTipText     =   "Edit Customer Information"
            Top             =   4260
            Width           =   3225
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Type"
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
            Height          =   465
            Index           =   12
            Left            =   120
            TabIndex        =   116
            Top             =   1590
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Instruction"
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
            Height          =   255
            Index           =   14
            Left            =   -780
            TabIndex        =   120
            Top             =   3540
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Delivery Address"
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
            Height          =   945
            Index           =   13
            Left            =   -120
            TabIndex        =   118
            Top             =   2130
            Width           =   2535
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "TIN"
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
            Left            =   90
            TabIndex        =   108
            Top             =   90
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Issued on"
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
            Left            =   90
            TabIndex        =   112
            Top             =   855
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Issued at"
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
            Index           =   11
            Left            =   90
            TabIndex        =   114
            Top             =   1245
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "CTC No."
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
            Index           =   9
            Left            =   90
            TabIndex        =   110
            Top             =   510
            Width           =   660
         End
      End
      Begin VB.PictureBox picVehiclesDetail 
         BorderStyle     =   0  'None
         Height          =   4635
         Left            =   -69970
         ScaleHeight     =   4635
         ScaleWidth      =   9630
         TabIndex        =   43
         Top             =   600
         Visible         =   0   'False
         Width           =   9630
         Begin VB.Frame fraPlateno 
            Caption         =   "PLATE NO"
            Enabled         =   0   'False
            Height          =   1470
            Left            =   4830
            TabIndex        =   45
            Top             =   60
            Width           =   4440
            Begin VB.TextBox txtVehicleWarrantyCertifcate 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1245
               TabIndex        =   51
               Top             =   975
               Width           =   3105
            End
            Begin VB.TextBox txtVehicleKMreading 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1230
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   600
               Width           =   3180
            End
            Begin MSMask.MaskEdBox txtVehiclePlateNo 
               Height          =   345
               Left            =   1230
               TabIndex        =   47
               Top             =   225
               Width           =   3180
               _ExtentX        =   5609
               _ExtentY        =   609
               _Version        =   393216
               BackColor       =   16777215
               ForeColor       =   7347754
               MaxLength       =   6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "dd-mmm-yy"
               PromptChar      =   "_"
            End
            Begin VB.Label Label71 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Warranty Certificate #"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   120
               TabIndex        =   50
               Top             =   960
               Width           =   960
            End
            Begin VB.Label Label72 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "KM Reading "
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
               Left            =   150
               TabIndex        =   48
               Top             =   630
               Width           =   1050
            End
            Begin VB.Label Label56 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Plate #"
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
               Left            =   525
               TabIndex        =   46
               Top             =   270
               Width           =   585
            End
         End
         Begin VB.CheckBox chkInsurance 
            Caption         =   "Vehicle is Insured"
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
            Height          =   255
            Left            =   4830
            TabIndex        =   84
            Top             =   2100
            Width           =   1815
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   4635
            Left            =   105
            ScaleHeight     =   4635
            ScaleWidth      =   4605
            TabIndex        =   52
            Top             =   435
            Width           =   4605
            Begin VB.TextBox txtVehicleTransmission 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   3510
               TabIndex        =   78
               Top             =   3840
               Width           =   945
            End
            Begin VB.TextBox txtVehicleModelCode 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   3270
               TabIndex        =   58
               ToolTipText     =   "Vehicles Make (Manufacturing Company)"
               Top             =   420
               Width           =   1200
            End
            Begin VB.TextBox txtVehicleColor 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1260
               TabIndex        =   77
               Top             =   3840
               Width           =   2205
            End
            Begin VB.TextBox txtVehicleDateReleased 
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
               Left            =   1260
               Locked          =   -1  'True
               TabIndex        =   75
               Text            =   "Text1"
               Top             =   3480
               Width           =   3225
            End
            Begin VB.TextBox txtVehicleProdNo 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1260
               TabIndex        =   65
               Tag             =   "@R"
               ToolTipText     =   "Vehicles Production Number"
               Top             =   1590
               Width           =   3225
            End
            Begin VB.TextBox txtVehicleEngineNo 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1260
               TabIndex        =   67
               Tag             =   "@R"
               ToolTipText     =   "Engine Number"
               Top             =   1950
               Width           =   3225
            End
            Begin VB.TextBox txtVehicleFrameNo 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1260
               TabIndex        =   69
               Top             =   2310
               Width           =   3225
            End
            Begin VB.TextBox txtVehicleDescription 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1260
               TabIndex        =   61
               ToolTipText     =   "Vehicles Production Number"
               Top             =   780
               Width           =   3225
            End
            Begin VB.TextBox txtVehicleConductionSticker 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1260
               TabIndex        =   63
               ToolTipText     =   "Vehicles Production Number"
               Top             =   1170
               Width           =   3225
            End
            Begin VB.TextBox txtVehicleModel 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1260
               TabIndex        =   59
               ToolTipText     =   "Vehicles Production Number"
               Top             =   420
               Width           =   1995
            End
            Begin VB.TextBox txtVehicleYear 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   3270
               TabIndex        =   56
               ToolTipText     =   "Vehicles Production Number"
               Top             =   30
               Width           =   1200
            End
            Begin VB.TextBox txtVehicleMake 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1260
               TabIndex        =   54
               ToolTipText     =   "Vehicles Make (Manufacturing Company)"
               Top             =   30
               Width           =   1515
            End
            Begin VB.TextBox txtVehicleSerialNo 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1260
               TabIndex        =   73
               Top             =   3090
               Width           =   3225
            End
            Begin VB.TextBox txtVehicleVinNo 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1260
               TabIndex        =   71
               Top             =   2700
               Width           =   3225
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Color :"
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
               Left            =   555
               TabIndex        =   76
               Top             =   3885
               Width           =   540
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Released Date"
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
               Left            =   0
               TabIndex        =   74
               Top             =   3480
               Width           =   1230
            End
            Begin VB.Label Label66 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
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
               Left            =   225
               TabIndex        =   60
               Top             =   765
               Width           =   975
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Frame No."
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
               Left            =   345
               TabIndex        =   68
               Top             =   2355
               Width           =   855
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Engine No."
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
               Left            =   315
               TabIndex        =   66
               Top             =   1995
               Width           =   885
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Make"
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
               Left            =   735
               TabIndex        =   53
               Top             =   105
               Width           =   465
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Prod. No."
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
               Index           =   0
               Left            =   435
               TabIndex        =   64
               Top             =   1635
               Width           =   765
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Conduction Sticker No"
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
               Height          =   405
               Left            =   90
               TabIndex        =   62
               Top             =   1095
               Width           =   1110
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Model/Code"
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
               Left            =   210
               TabIndex        =   57
               Top             =   405
               Width           =   990
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00400000&
               Height          =   225
               Left            =   2835
               TabIndex        =   55
               Top             =   105
               Width           =   390
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Serial No."
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
               Left            =   390
               TabIndex        =   72
               Top             =   3105
               Width           =   810
            End
            Begin VB.Label Label57 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Vin No"
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
               Left            =   660
               TabIndex        =   70
               Top             =   2730
               Width           =   540
            End
         End
         Begin VB.CommandButton cmdSelectVehicles 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Select Vehicles"
            Height          =   315
            Left            =   1365
            MaskColor       =   &H00400000&
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   75
            Width           =   3195
         End
         Begin VB.PictureBox picInsurance 
            BorderStyle     =   0  'None
            Height          =   1545
            Left            =   6570
            ScaleHeight     =   1545
            ScaleWidth      =   3555
            TabIndex        =   79
            Top             =   2070
            Width           =   3555
            Begin VB.ComboBox cboInsuranceCompany 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00400000&
               Height          =   345
               ItemData        =   "VehicleInvoice.frx":8907
               Left            =   120
               List            =   "VehicleInvoice.frx":8909
               TabIndex        =   83
               Top             =   960
               Width           =   2505
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   345
               Left            =   120
               TabIndex        =   81
               Top             =   270
               Width           =   2535
               _ExtentX        =   4471
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
               CalendarTitleBackColor=   8388608
               CalendarTitleForeColor=   16777215
               CheckBox        =   -1  'True
               Format          =   54329345
               CurrentDate     =   39213
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Insurance Company"
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
               Left            =   120
               TabIndex        =   82
               Top             =   690
               Width           =   1695
            End
            Begin VB.Label Label70 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Insured Date"
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
               Left            =   90
               TabIndex        =   80
               Top             =   30
               Width           =   1080
            End
         End
         Begin VB.Label lblVehicleStatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Left            =   4800
            TabIndex        =   85
            Top             =   3930
            Width           =   4545
         End
      End
      Begin VB.PictureBox picTerms 
         BorderStyle     =   0  'None
         Height          =   5265
         Left            =   -69970
         ScaleHeight     =   5265
         ScaleWidth      =   9855
         TabIndex        =   123
         Top             =   600
         Visible         =   0   'False
         Width           =   9855
         Begin VB.PictureBox fraTermsCredit 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   5220
            Left            =   30
            ScaleHeight     =   5220
            ScaleWidth      =   9780
            TabIndex        =   124
            Top             =   0
            Width           =   9780
            Begin VB.TextBox txtFinNoOfTermAmort 
               Alignment       =   1  'Right Justify
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
               Height          =   360
               Left            =   1830
               MaxLength       =   10
               TabIndex        =   143
               TabStop         =   0   'False
               Top             =   4200
               Visible         =   0   'False
               Width           =   2775
            End
            Begin VB.Frame Frame2 
               BorderStyle     =   0  'None
               Height          =   5205
               Left            =   4980
               TabIndex        =   144
               Top             =   0
               Width           =   4785
               Begin VB.TextBox LAB_TOTAL_FIN 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00400000&
                  Height          =   405
                  Left            =   2070
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   170
                  TabStop         =   0   'False
                  Top             =   4800
                  Width           =   2610
               End
               Begin VB.CheckBox chkZeroRate2 
                  Caption         =   "Zero Rated Tax"
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
                  Height          =   435
                  Left            =   480
                  TabIndex        =   151
                  Top             =   1410
                  Width           =   1155
               End
               Begin VB.TextBox txtFinTax 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   2100
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   153
                  Top             =   1440
                  Width           =   2625
               End
               Begin VB.TextBox txtFinDownpaymentRate 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   390
                  Left            =   2100
                  MaxLength       =   10
                  TabIndex        =   156
                  Top             =   1845
                  Width           =   780
               End
               Begin VB.TextBox txtFinChattel 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   390
                  Left            =   2085
                  MaxLength       =   20
                  TabIndex        =   162
                  Top             =   3105
                  Width           =   2625
               End
               Begin VB.TextBox txtFinAccessories 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
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
                  ForeColor       =   &H00400000&
                  Height          =   390
                  Left            =   2085
                  MaxLength       =   20
                  TabIndex        =   164
                  Top             =   3510
                  Width           =   2625
               End
               Begin VB.TextBox txtFinFreight 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   390
                  Left            =   2085
                  MaxLength       =   20
                  TabIndex        =   166
                  Top             =   3930
                  Width           =   2625
               End
               Begin VB.TextBox txtFinOthers 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   390
                  Left            =   2085
                  MaxLength       =   20
                  TabIndex        =   168
                  Top             =   4350
                  Width           =   2625
               End
               Begin VB.TextBox txtFinInsurance 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   390
                  Left            =   2085
                  MaxLength       =   20
                  TabIndex        =   158
                  Top             =   2265
                  Width           =   2625
               End
               Begin VB.TextBox txtFinLTORegFee 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   390
                  Left            =   2085
                  MaxLength       =   20
                  TabIndex        =   160
                  Top             =   2685
                  Width           =   2625
               End
               Begin VB.TextBox txtFinDownPayment 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   390
                  Left            =   2910
                  MaxLength       =   20
                  TabIndex        =   155
                  Top             =   1830
                  Width           =   1800
               End
               Begin VB.TextBox txtFinDiscount 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   390
                  Left            =   2100
                  MaxLength       =   20
                  TabIndex        =   148
                  Top             =   585
                  Width           =   2625
               End
               Begin VB.TextBox txtFinSalesPrice 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   390
                  Left            =   2100
                  MaxLength       =   20
                  TabIndex        =   146
                  Top             =   150
                  Width           =   2625
               End
               Begin VB.TextBox txtFinOthersDesc 
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00701E2A&
                  Height          =   390
                  Left            =   90
                  MaxLength       =   10
                  TabIndex        =   167
                  Text            =   " "
                  Top             =   4350
                  Width           =   1950
               End
               Begin VB.TextBox txtFinNetSalesPrice 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   390
                  Left            =   2115
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   150
                  Top             =   1005
                  Width           =   2595
               End
               Begin VB.Label Label34 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TOTAL AMOUNT DUE"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   240
                  Left            =   60
                  TabIndex        =   169
                  Top             =   4890
                  Width           =   1995
               End
               Begin VB.Label Label24 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "TAX"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00400000&
                  Height          =   210
                  Index           =   13
                  Left            =   1665
                  TabIndex        =   152
                  Top             =   1500
                  Width           =   330
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "FREIGHT && HANDLING : "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00400000&
                  Height          =   210
                  Index           =   8
                  Left            =   225
                  TabIndex        =   165
                  Top             =   4005
                  Width           =   1815
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Discount"
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
                  Index           =   11
                  Left            =   1275
                  TabIndex        =   147
                  Top             =   645
                  Width           =   750
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "CHMO Fee"
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
                  Index           =   17
                  Left            =   1125
                  TabIndex        =   161
                  Top             =   3135
                  Width           =   870
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "LTO Reg. Fee"
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
                  Index           =   16
                  Left            =   885
                  TabIndex        =   159
                  Top             =   2730
                  Width           =   1110
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Insurance"
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
                  Index           =   15
                  Left            =   1140
                  TabIndex        =   157
                  Top             =   2340
                  Width           =   855
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
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
                  Index           =   14
                  Left            =   735
                  TabIndex        =   154
                  Top             =   1905
                  Width           =   1275
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sales Price"
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
                  Left            =   1050
                  TabIndex        =   145
                  Top             =   225
                  Width           =   975
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "FREEBIES"
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
                  Index           =   18
                  Left            =   1185
                  TabIndex        =   163
                  Top             =   3555
                  Width           =   810
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Net Sales Price"
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
                  Index           =   12
                  Left            =   750
                  TabIndex        =   149
                  Top             =   1110
                  Width           =   1305
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Financing Details"
               Height          =   3465
               Left            =   90
               TabIndex        =   125
               Top             =   30
               Width           =   4680
               Begin VB.CheckBox cmdAuto 
                  Caption         =   "Auto Compute"
                  Height          =   315
                  Left            =   1770
                  Style           =   1  'Graphical
                  TabIndex        =   270
                  Top             =   1470
                  Width           =   2775
               End
               Begin VB.TextBox txtFinBankTerm 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   1770
                  MaxLength       =   10
                  TabIndex        =   131
                  Top             =   690
                  Width           =   2775
               End
               Begin VB.TextBox txtFinBaltoFinanced 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
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
                  Height          =   360
                  Left            =   1770
                  Locked          =   -1  'True
                  TabIndex        =   137
                  TabStop         =   0   'False
                  Top             =   2220
                  Width           =   2775
               End
               Begin VB.TextBox txtFinNetMonthlyAmort 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   1770
                  TabIndex        =   135
                  TabStop         =   0   'False
                  Top             =   1830
                  Width           =   2775
               End
               Begin VB.TextBox txtFinAOR 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   1770
                  MaxLength       =   10
                  TabIndex        =   133
                  Top             =   1080
                  Width           =   2775
               End
               Begin VB.TextBox txtFinGMI 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   1770
                  MaxLength       =   10
                  TabIndex        =   139
                  Top             =   2595
                  Width           =   2775
               End
               Begin VB.TextBox txtFinRPPD 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   1770
                  MaxLength       =   10
                  TabIndex        =   141
                  Top             =   2985
                  Width           =   2775
               End
               Begin VB.ComboBox cboFinFinancingCo 
                  Appearance      =   0  'Flat
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
                  ForeColor       =   &H00400000&
                  Height          =   345
                  Left            =   1770
                  TabIndex        =   129
                  Top             =   315
                  Width           =   2775
               End
               Begin VB.TextBox txtFinModeofPayment 
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   1770
                  MaxLength       =   10
                  TabIndex        =   127
                  Top             =   -375
                  Visible         =   0   'False
                  Width           =   2775
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
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
                  Index           =   4
                  Left            =   1320
                  TabIndex        =   132
                  Top             =   1110
                  Width           =   375
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bank Terms"
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
                  Index           =   3
                  Left            =   645
                  TabIndex        =   130
                  Top             =   750
                  Width           =   1035
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
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
                  Index           =   6
                  Left            =   150
                  TabIndex        =   136
                  Top             =   2280
                  Width           =   1560
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "GMI"
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
                  Left            =   1395
                  TabIndex        =   138
                  Top             =   2625
                  Width           =   315
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "RPPD"
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
                  Left            =   1230
                  TabIndex        =   140
                  Top             =   3075
                  Width           =   480
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
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
                  Index           =   5
                  Left            =   465
                  TabIndex        =   134
                  Top             =   1920
                  Width           =   1245
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Financing Co."
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
                  Index           =   2
                  Left            =   585
                  TabIndex        =   128
                  Top             =   390
                  Width           =   1125
               End
               Begin VB.Label Label76 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mode Of Payment"
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
                  Left            =   150
                  TabIndex        =   126
                  Top             =   -300
                  Visible         =   0   'False
                  Width           =   1500
               End
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "No of Terms"
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
               Index           =   9
               Left            =   1830
               TabIndex        =   142
               Top             =   3960
               Visible         =   0   'False
               Width           =   1035
            End
         End
         Begin VB.PictureBox fraTermsCash 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   5220
            Left            =   30
            ScaleHeight     =   5220
            ScaleWidth      =   9840
            TabIndex        =   171
            Top             =   0
            Width           =   9840
            Begin VB.Frame Frame4 
               BorderStyle     =   0  'None
               Height          =   4725
               Left            =   4920
               TabIndex        =   174
               Top             =   30
               Width           =   4875
               Begin VB.TextBox LAB_TOTAL_CASH 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00400000&
                  Height          =   405
                  Left            =   2190
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   195
                  TabStop         =   0   'False
                  Top             =   3930
                  Width           =   2535
               End
               Begin VB.CheckBox chkZeroRate1 
                  Caption         =   "Zero Rated Tax"
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
                  Height          =   435
                  Left            =   390
                  TabIndex        =   193
                  Top             =   3540
                  Width           =   1155
               End
               Begin VB.TextBox txtCashTax 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   2190
                  MaxLength       =   20
                  TabIndex        =   192
                  Top             =   3495
                  Width           =   2625
               End
               Begin VB.TextBox txtCashOthers 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   2190
                  MaxLength       =   20
                  TabIndex        =   190
                  Top             =   3075
                  Width           =   2625
               End
               Begin VB.TextBox txtCashSalesPrice 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   2190
                  MaxLength       =   20
                  TabIndex        =   176
                  Top             =   180
                  Width           =   2625
               End
               Begin VB.TextBox txtCashNetSalesPrice 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  ForeColor       =   &H00400000&
                  Height          =   360
                  Left            =   2190
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   180
                  Top             =   975
                  Width           =   2625
               End
               Begin VB.TextBox txtCashInsurance 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   2190
                  MaxLength       =   20
                  TabIndex        =   182
                  Top             =   1425
                  Width           =   2625
               End
               Begin VB.TextBox txtCashLTORegFee 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   2190
                  MaxLength       =   20
                  TabIndex        =   184
                  Top             =   1845
                  Width           =   2625
               End
               Begin VB.TextBox txtCashAccessories 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
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
                  ForeColor       =   &H00400000&
                  Height          =   360
                  Left            =   2190
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   186
                  Top             =   2250
                  Width           =   2625
               End
               Begin VB.TextBox txtCashDiscount 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   2190
                  MaxLength       =   20
                  TabIndex        =   178
                  Top             =   570
                  Width           =   2625
               End
               Begin VB.TextBox txtCashFreight 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   360
                  Left            =   2190
                  MaxLength       =   20
                  TabIndex        =   188
                  Top             =   2670
                  Width           =   2625
               End
               Begin VB.TextBox txtCashOthersDesc 
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
                  ForeColor       =   &H00400000&
                  Height          =   360
                  Left            =   330
                  TabIndex        =   189
                  Text            =   " "
                  Top             =   3105
                  Width           =   1800
               End
               Begin VB.Label Label75 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "TOTAL AMOUNT DUE"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   240
                  Left            =   30
                  TabIndex        =   194
                  Top             =   4050
                  Width           =   1995
               End
               Begin VB.Label Label61 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "LTO Reg. Fee"
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
                  Left            =   945
                  TabIndex        =   183
                  Top             =   1920
                  Width           =   1110
               End
               Begin VB.Label Label40 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Discount"
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
                  Left            =   1305
                  TabIndex        =   177
                  Top             =   645
                  Width           =   750
               End
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Insurance"
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
                  Left            =   1200
                  TabIndex        =   181
                  Top             =   1470
                  Width           =   855
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sales Price"
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
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   175
                  Top             =   255
                  Width           =   975
               End
               Begin VB.Label Label25 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Net Sales Price"
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
                  Left            =   750
                  TabIndex        =   179
                  Top             =   1035
                  Width           =   1305
               End
               Begin VB.Label Label52 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "FREE BEEIES"
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
                  Left            =   975
                  TabIndex        =   185
                  Top             =   2325
                  Width           =   1065
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "FREIGHT && HANDLING "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00400000&
                  Height          =   210
                  Index           =   1
                  Left            =   435
                  TabIndex        =   187
                  Top             =   2745
                  Width           =   1725
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "TAX : "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00400000&
                  Height          =   210
                  Index           =   2
                  Left            =   1590
                  TabIndex        =   191
                  Top             =   3555
                  Width           =   465
               End
            End
            Begin VB.ComboBox cboCashModeofPayment 
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
               Left            =   240
               TabIndex        =   173
               Top             =   330
               Width           =   2235
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Mode Of Payment"
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
               Left            =   225
               TabIndex        =   172
               Top             =   75
               Width           =   1500
            End
         End
      End
   End
   Begin VB.PictureBox picMultipleSO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   1920
      ScaleHeight     =   4965
      ScaleWidth      =   6735
      TabIndex        =   196
      Top             =   1980
      Visible         =   0   'False
      Width           =   6765
      Begin VB.ComboBox cboMultiCustomer 
         Height          =   330
         Left            =   780
         TabIndex        =   199
         Text            =   "cboMultiCustomer"
         ToolTipText     =   "Search By Customer"
         Top             =   420
         Width           =   5775
      End
      Begin VB.CommandButton cmdCloseMultiple 
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
         Left            =   6390
         TabIndex        =   198
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmdCancelMultiple 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   5730
         TabIndex        =   203
         ToolTipText     =   "Cancel"
         Top             =   4410
         Width           =   825
      End
      Begin VB.CommandButton cmdSelectMultiple 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4920
         TabIndex        =   202
         ToolTipText     =   "Select"
         Top             =   4410
         Width           =   825
      End
      Begin MSComctlLib.ListView lstMultipleSO 
         Height          =   3585
         Left            =   90
         TabIndex        =   201
         Top             =   810
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   6324
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SO_NO"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CUSTNAME"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "MODELDESCRIPTION"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "IGNKEY_NO"
            Object.Width           =   1658
         EndProperty
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "FILTER"
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
         Height          =   285
         Index           =   1
         Left            =   60
         TabIndex        =   200
         Top             =   450
         Width           =   2505
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   285
         Left            =   -15
         TabIndex        =   197
         Top             =   0
         Width           =   6915
         _Version        =   655364
         _ExtentX        =   12197
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Select Vehicle Details"
         ForeColor       =   -2147483630
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
         ForeColor       =   -2147483630
      End
   End
   Begin VB.PictureBox picViewVehicles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4950
      Left            =   150
      ScaleHeight     =   4920
      ScaleWidth      =   9720
      TabIndex        =   204
      Top             =   2220
      Visible         =   0   'False
      Width           =   9750
      Begin XtremeReportControl.ReportControl lvViewVehicles 
         Height          =   3195
         Left            =   60
         TabIndex        =   210
         Top             =   750
         Width           =   9540
         _Version        =   655364
         _ExtentX        =   16828
         _ExtentY        =   5636
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Detail Information"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   207
         Top             =   420
         Width           =   1665
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "&Cancel"
         Height          =   735
         Index           =   0
         Left            =   8850
         MouseIcon       =   "VehicleInvoice.frx":890B
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":8A5D
         Style           =   1  'Graphical
         TabIndex        =   213
         ToolTipText     =   "Cancel"
         Top             =   3990
         Width           =   705
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
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
         Left            =   9345
         TabIndex        =   206
         Top             =   15
         Width           =   285
      End
      Begin VB.TextBox txtFilterViewVehicles 
         Height          =   375
         Left            =   5460
         TabIndex        =   208
         Top             =   375
         Width           =   4155
      End
      Begin VB.CommandButton cmdSelectViewVehicles 
         Caption         =   "&Select"
         Height          =   735
         Left            =   8160
         MouseIcon       =   "VehicleInvoice.frx":8D9B
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":8EED
         Style           =   1  'Graphical
         TabIndex        =   212
         ToolTipText     =   "Select"
         Top             =   3990
         Width           =   705
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "F# = Frame No . V#= VIN No .S#=Serial No"
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   75
         TabIndex        =   214
         Top             =   4560
         Width           =   7515
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
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
         Height          =   285
         Index           =   10
         Left            =   4710
         TabIndex        =   209
         Top             =   450
         Width           =   2505
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Left            =   0
         TabIndex        =   205
         Top             =   0
         Width           =   9720
         _Version        =   655364
         _ExtentX        =   17145
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "::: Vehicles Inventory:::"
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
         VisualTheme     =   3
         Alignment       =   1
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "C#= Conduction Sticker No . P#= Production No. E# = Engine No ."
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   75
         TabIndex        =   211
         Top             =   4335
         Width           =   5175
      End
   End
   Begin VB.PictureBox picNetSpeed 
      Height          =   1935
      Left            =   3600
      ScaleHeight     =   1875
      ScaleWidth      =   2625
      TabIndex        =   276
      Top             =   3180
      Visible         =   0   'False
      Width           =   2685
      Begin VB.TextBox txtOldCS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1350
         TabIndex        =   279
         Top             =   1455
         Width           =   1215
      End
      Begin VB.TextBox txtMCode 
         BackColor       =   &H00FF0000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1365
         TabIndex        =   278
         TabStop         =   0   'False
         Top             =   30
         Width           =   1215
      End
      Begin VB.TextBox labProspectID 
         BackColor       =   &H00FF0000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1365
         TabIndex        =   277
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "OLD CS"
         Height          =   375
         Index           =   4
         Left            =   180
         TabIndex        =   286
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "prospectid"
         Height          =   375
         Index           =   3
         Left            =   270
         TabIndex        =   285
         Top             =   1110
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "labeditdetail"
         Height          =   375
         Index           =   2
         Left            =   300
         TabIndex        =   284
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "labid"
         Height          =   375
         Index           =   1
         Left            =   300
         TabIndex        =   283
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "MCODE"
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   282
         Top             =   150
         Width           =   975
      End
      Begin VB.Label labid 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1365
         TabIndex        =   281
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label labEDITDetail 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FALSE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1365
         TabIndex        =   280
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.PictureBox picCancelReason 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   2910
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   2955
      ScaleWidth      =   4920
      TabIndex        =   230
      Top             =   2880
      Visible         =   0   'False
      Width           =   4950
      Begin VB.CommandButton cmdCancelReason 
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
         Index           =   3
         Left            =   4500
         TabIndex        =   236
         Top             =   30
         Width           =   285
      End
      Begin VB.TextBox txtReasonCancel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1485
         Left            =   270
         MaxLength       =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   233
         Top             =   810
         Width           =   4410
      End
      Begin VB.CommandButton cmdCancelReason 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3750
         TabIndex        =   235
         ToolTipText     =   "Cancel"
         Top             =   2400
         Width           =   945
      End
      Begin VB.CommandButton cmdCancelFinal 
         Caption         =   "Confirm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2820
         TabIndex        =   234
         ToolTipText     =   "Confirm Reason"
         Top             =   2400
         Width           =   945
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   330
         Left            =   0
         TabIndex        =   231
         Top             =   0
         Width           =   4935
         _Version        =   655364
         _ExtentX        =   8705
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   ":: Input Reason ::"
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
      Begin VB.Label Label68 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "State Reason Of Cancelation of This Invoice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   240
         TabIndex        =   232
         Top             =   555
         Width           =   3690
      End
   End
   Begin VB.PictureBox picAccessories 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   3090
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   3105
      ScaleWidth      =   4920
      TabIndex        =   215
      Top             =   2970
      Visible         =   0   'False
      Width           =   4950
      Begin VB.CheckBox chISFREE 
         BackColor       =   &H00FFFFFF&
         Caption         =   "With Charge"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   435
         Left            =   1350
         TabIndex        =   288
         Top             =   2190
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdCancelDetailProduct 
         Height          =   795
         Index           =   0
         Left            =   4020
         MouseIcon       =   "VehicleInvoice.frx":9229
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":937B
         Style           =   1  'Graphical
         TabIndex        =   227
         TabStop         =   0   'False
         ToolTipText     =   "Exit Entry"
         Top             =   2250
         Width           =   555
      End
      Begin VB.ComboBox cboAccessories 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1350
         TabIndex        =   218
         Top             =   675
         Width           =   3270
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
         Height          =   315
         Index           =   1
         Left            =   4620
         TabIndex        =   228
         Top             =   0
         Width           =   315
      End
      Begin VB.TextBox txtAccQty 
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
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   223
         Text            =   "1"
         Top             =   1470
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtAccRate 
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
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   220
         Top             =   1050
         Width           =   2055
      End
      Begin VB.TextBox txtAccAmount 
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
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   225
         TabStop         =   0   'False
         Top             =   1815
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdOkMaterials 
         Height          =   795
         Left            =   3480
         MouseIcon       =   "VehicleInvoice.frx":96E1
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":9833
         Style           =   1  'Graphical
         TabIndex        =   229
         TabStop         =   0   'False
         ToolTipText     =   "Save Entry"
         Top             =   2250
         Width           =   555
      End
      Begin VB.CommandButton Command5 
         Height          =   795
         Left            =   2940
         MouseIcon       =   "VehicleInvoice.frx":9B83
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":9CD5
         Style           =   1  'Graphical
         TabIndex        =   226
         TabStop         =   0   'False
         ToolTipText     =   "Delete Entry"
         Top             =   2250
         Width           =   555
      End
      Begin VB.Label labAccID 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   150
         TabIndex        =   222
         Top             =   1575
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
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
         Index           =   2
         Left            =   1740
         TabIndex        =   224
         Top             =   1905
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
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
         Index           =   1
         Left            =   1695
         TabIndex        =   221
         Top             =   1560
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
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
         Index           =   3
         Left            =   1770
         TabIndex        =   219
         Top             =   1140
         Width           =   705
      End
      Begin VB.Label Label64 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   225
         TabIndex        =   217
         Top             =   750
         Width           =   975
      End
      Begin XtremeShortcutBar.ShortcutCaption capAccessories 
         Height          =   330
         Left            =   0
         TabIndex        =   216
         Top             =   0
         Width           =   4935
         _Version        =   655364
         _ExtentX        =   8705
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "Free Beeies"
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
         GradientColorLight=   12632256
         GradientColorDark=   8421504
         ForeColor       =   -2147483630
      End
   End
   Begin VB.PictureBox picRelease 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2550
      Left            =   3390
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   2520
      ScaleWidth      =   4080
      TabIndex        =   237
      Top             =   3120
      Visible         =   0   'False
      Width           =   4110
      Begin MSComCtl2.DTPicker txtRelease_Time 
         Height          =   375
         Left            =   1680
         TabIndex        =   246
         Top             =   1290
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   12582912
         CalendarTitleBackColor=   12582912
         CalendarTitleForeColor=   12582912
         CalendarTrailingForeColor=   12582912
         Format          =   54329346
         CurrentDate     =   39506
      End
      Begin VB.CommandButton cmdCancelRelease 
         CausesValidation=   0   'False
         Height          =   645
         Index           =   0
         Left            =   3060
         MouseIcon       =   "VehicleInvoice.frx":A000
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":A152
         Style           =   1  'Graphical
         TabIndex        =   239
         TabStop         =   0   'False
         ToolTipText     =   "Cancel Entry"
         Top             =   1770
         Width           =   675
      End
      Begin VB.TextBox txtRelease_VDR 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   405
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   242
         Top             =   420
         Width           =   2055
      End
      Begin VB.TextBox txtRelease_Date 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   243
         Top             =   870
         Width           =   2055
      End
      Begin VB.CommandButton cmdCancelRelease 
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
         Index           =   5
         Left            =   3750
         TabIndex        =   240
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.CommandButton cmdReleaseVehicle 
         Height          =   645
         Left            =   2340
         MouseIcon       =   "VehicleInvoice.frx":A490
         MousePointer    =   99  'Custom
         Picture         =   "VehicleInvoice.frx":A5E2
         Style           =   1  'Graphical
         TabIndex        =   247
         ToolTipText     =   "Save Entry"
         Top             =   1770
         Width           =   735
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Time Release:"
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
         Index           =   0
         Left            =   360
         TabIndex        =   244
         Top             =   1380
         Width           =   1200
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VDR NO#"
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
         Left            =   780
         TabIndex        =   241
         Top             =   465
         Width           =   765
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   330
         Left            =   0
         TabIndex        =   238
         Top             =   0
         Width           =   4110
         _Version        =   655364
         _ExtentX        =   7250
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "Released Vehicle"
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
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Released:"
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
         Left            =   330
         TabIndex        =   245
         Top             =   960
         Width           =   1275
      End
   End
   Begin VB.Label labStatus 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   7260
      TabIndex        =   31
      Top             =   1620
      Width           =   2625
   End
   Begin VB.Label lblVehicleInformation 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   180
      TabIndex        =   32
      Top             =   1620
      Width           =   7080
   End
End
Attribute VB_Name = "frmSMIS_Trans_VehicleInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsInvoice                                                         As ADODB.Recordset
Dim rsPurchAgreeClone                                                 As ADODB.Recordset
Dim rsMRRINV                                                          As ADODB.Recordset
Attribute rsMRRINV.VB_VarUserMemId = 1073938435
Dim rsFinCom                                                          As ADODB.Recordset
Dim rsSignatories                                                     As ADODB.Recordset
Dim MayoModel                                                         As Boolean
Dim AddorEdit                                                         As String
Dim Tutal                                                             As Double
Private WithEvents frmCustomerInfo                                    As frmAllCustomer
Attribute frmCustomerInfo.VB_VarHelpID = -1
Dim PROSPECTID                                                        As Long
Dim MULTIPLEVI                                                        As Boolean
Dim ComputebyPert                                                     As Boolean

Function GenerateID(FLDNAME As String) As String
    Dim rsID                                                          As ADODB.Recordset

    Set rsID = gconDMIS.Execute("Select MAX(" & FLDNAME & " ) as SO_NO from SMIS_SalesOrder")
    If rsID.Fields(0).Value = 0 Then
        GenerateID = Format(1, "00000000")
    Else
        GenerateID = Format(Val(N2Str2Zero(rsID![SO_NO])) + 1, "00000000")

    End If
    Set rsID = Nothing

End Function




'"update SMIS_SALESORDER set" & _
 " code = " & vtxtCusCode & ", deyt = " & vtxtDate & "," & _
 " hometelno = " & vTxtHomeTelNo & "," & _
 " officeadd = " & vTxtOfficeAdd & ", officetelno = " & vTxtOfficeTelNo & "," & _
 " birthdate = " & vTxtBirthDate & ", spouse = " & vTxtSpouse & "," & _
 " person = " & vTxtPerson & ", posisyon = " & vTxtPerson & "," & _
 " tin = " & vTxtTIN & ", ctcno = " & vTxtCTCNo & "," & _
 " issuedat = " & vTxtIssuedAt & ", issuedon = " & vTxtIssuedOn & "," & _
 " model = " & N2Str2Null(txtVehicleModel) & ", prodno = '" & txtVehicleProdNo & "'," & _
 " engineno = " & vTxtEngineNo & ", ignkey_no = " & vtxtConductionStickerNo & ", frameno = " & vTxtFrameNo & "," & _
 " color = " & vcboColor & ", type = " & vcboType & "," & _
 " term = '" & TIRM & "', financingco = " & vcboFinancingCo & "," & _
 " salesae = " & vcboSalesAE & ", salesprice =" & N2Str2Zero(txtCashSalesPrice) & ", netsalesprice = " & N2Str2Zero(txtCashNetSalesPrice) & "," & _
 " insurance = " & N2Str2Zero(txtCashInsurance) & ", LTOREGFee = " & NumericVal(txtCashLTORegFee) & ", accessories = " & N2Str2Zero(txtCashAccessories) & ", tax = " & N2Str2Zero(txtCashTax) & ", others = " & N2Str2Zero(txtCashOthers) & ",additionalinfo = " & vtxtCashAdditionalInfo & "," & _
 " total = " & N2Str2Zero(txtCashTotal) & ", vi_no = " & vtxtVINo & "," & _
 " certific8= " & N2Str2Null(txtVehicleWarrantyCertifcate) & "," & _
 " vdr_no = " & vtxtVDRNo & ", plate_no = " & vtxtPlate_No & ", preparedby = " & vtxtPreparedBy & ", checkedby = " & vtxtCheckedBy & "," & _
 " salesapproved = " & vtxtSalesApproved & ", salesdispatcher = " & vtxtSalesDispatcher & ", bankterm = " & vcboBankTerm & ", datereleased = " & vtxtDateReleased & ", insured = '" & INSURE & "', ModeOfPayment = " & vtxtModeOfPayment & ", DownpaymentRate = " & vtxtDownpaymentRate & ", Terms = " & vtxtTerms & _
 " where id = " & labid.caption

Function GetAccountType(XXX)
    If XXX = "F" Then
        GetAccountType = "Fleet"
    ElseIf XXX = "R" Then
        GetAccountType = "Retail"
    Else
        GetAccountType = ""
    End If
End Function

Function GetPo(XXX)
    If XXX = "CPO" Then
        GetPo = "Company PO"
    ElseIf XXX = "CSH" Or XXX = "COD" Then
        GetPo = "Cash"
    ElseIf XXX = "CHK" Then
        GetPo = "Cheque"
    End If
End Function

Function SetAccountType(XXX)
    If UCase(XXX) = "FLEET" Then
        SetAccountType = "F"
    ElseIf UCase(XXX) = "RETAIL" Then
        SetAccountType = "R"
    Else
        SetAccountType = ""
    End If
End Function

Function SetColor(CCC As String)
    Dim rsColor                                                       As ADODB.Recordset
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select COLOR_CODE,COLOR_DESC from ALL_Color where COLOR_DESC = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        SetColor = Null2String(rsColor!Color_code)
    Else
        SetColor = ""
    End If
End Function

Function setFinCode(fff As String)
    Set rsFinCom = New ADODB.Recordset
    rsFinCom.Open "select * from SMIS_FinCom where company = '" & fff & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsFinCom.EOF And Not rsFinCom.BOF Then
        setFinCode = Null2String(rsFinCom!CODE)
    Else
        MsgSpeechBox "Invalid Financing Company ..." & vbCrLf & _
                     "Financing company must be added in Master File."
        setFinCode = ""
    End If
End Function

Function setFinCom(CCC As String)
    Set rsFinCom = New ADODB.Recordset
    rsFinCom.Open "select * from SMIS_FinCom where code = '" & CCC & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsFinCom.EOF And Not rsFinCom.BOF Then
        setFinCom = Null2String(rsFinCom!company)
    Else
        setFinCom = ""
    End If
End Function

Function SetMRRCode(CCC As String)
    Set rsMRRINV = New ADODB.Recordset
    rsMRRINV.Open "select * from SMIS_MRRINV_TABLE WHERE prodno = '" & CCC & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
        SetMRRCode = Null2String(rsMRRINV!CODE)
    Else
        MayoModel = True
    End If
End Function

Function SetPo(XXX)
    'If xxx = "Company PO" Then
    '    SetPo = "CPO"
    If XXX = "Cash" Then
        SetPo = "CSH"
    ElseIf XXX = "Cheque" Then
        SetPo = "CHK"
    End If

End Function

Function GetModelCode(XXXDESCRIPT As String) As String
    Dim RSMODELX                                                      As ADODB.Recordset
    Set RSMODELX = gconDMIS.Execute("SELECT CODE  FROM ALL_MODELCODE WHERE DESCRIPTION =" & N2Str2Null(XXXDESCRIPT))
    If Not (RSMODELX.EOF Or RSMODELX.BOF) Then
        GetModelCode = Null2String(RSMODELX!CODE)
    End If
    Set RSMODELX = Nothing
End Function

Private Function AORVALUE(Principal, AOR, TERM) As Double
    'On Error Resume Next

    If AOR <= 0 Then Exit Function
    If Principal <= 0 Then Exit Function
    If TERM <= 0 Then Exit Function
    Dim Interest                                                      As Double
    '    Interest = NumericVal(AOR)
    '    Interest = AOR / 1200
    '    AORVALUE = FormatNumber((Principal * Interest / (1 - ((1 / (1 + Interest) ^ Term)))), 2)
    AORVALUE = FormatNumber((NumericVal(txtFinBaltoFinanced) * (1 + (AOR / 100))) / TERM, 2)

End Function

Sub FillCboBankTerm()
    Set rsPurchAgreeClone = New ADODB.Recordset
    rsPurchAgreeClone.Open "select * from SMIS_PurchAgree", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPurchAgreeClone.EOF And Not rsPurchAgreeClone.BOF Then
        rsPurchAgreeClone.MoveFirst
        txtFinBankTerm.Clear
        Do While Not rsPurchAgreeClone.EOF
            txtFinBankTerm.AddItem Null2String(rsPurchAgreeClone!BANKTERM)
            rsPurchAgreeClone.MoveNext
        Loop
    Else
        txtFinBankTerm = ""
    End If
    Set rsPurchAgreeClone = Nothing
End Sub

Sub initMemvars()
    dtDateInvoiced.Value = FormatDateTime(LOGDATE, vbShortDate)
    labSJ = "": labDetails = "": labORNo = ""
    txtOldCS = "": txtMCode = "": labid = 0: labProspectID = 0: labEDITDetail = ""

    cmdAuto.Value = 0

    chkZeroRate2.Value = 0
    chkZeroRate1 = 0
    txtSODate = ""
    txtCustName = ""
    txtHomeAdd = ""
    txtTelephoneHome = ""
    txtOfficeAdd = ""
    txtTelephoneOffice = ""
    txtDateBirth = ""
    txtSpouse = ""
    txtContactPerson = ""
    txtPosition = ""
    txtTin = ""
    txtCTCNo = ""
    txtIssuedAt = ""
    txtIssuedOn = ""
    txtDeliveryAddress = ""
    txtVehicleProdNo = ""
    txtVehicleEngineNo = ""
    txtVehicleFrameNo = ""
    txtDeliveryInstruction = ""
    cboAccountType = ""

    txtFinBankTerm = ""

    txtCashSalesPrice = "0.00"
    txtCashNetSalesPrice = "0.00"
    txtCashInsurance = "0.00"
    txtCashAccessories = "0.00"
    txtCashOthers = "0.00"
    txtCashDiscount = "0.00"
    txtCashLTORegFee = "0.00"
    txtCashTax = "0.00"
    txtFinTax = "0.00"
    txtCashFreight = "0.00"
    txtCashOthersDesc = ""
    LAB_TOTAL_CASH = "0.00"

    txtFinNetSalesPrice = "0.00"
    txtFinDownPayment = "0.00"
    txtFinSalesPrice = "0.00"
    txtFinBaltoFinanced = "0.00"
    txtFinInsurance = "0.00"
    txtFinLTORegFee = "0.00"
    txtFinChattel = "0.00"
    txtFinAccessories = "0.00"
    txtFinOthers = "0.00"
    LAB_TOTAL_FIN = "0.00"

    infoAdditionalInfo = ""
    txtFinGMI = "0.00"
    txtFinRPPD = "0.00"
    txtFinNoOfTermAmort = "0.00"
    txtFinNetMonthlyAmort = "0.00"

    txtVehicleConductionSticker = ""
    labVDRNo = ""
    txtVINO = ""
    txtVehiclePlateNo = ""

    txtVehicleDateReleased = ""


    txtPreparedBy = ""
    txtCheckedBy = ""
    txtSalesApproved = ""

    txtGeneralManager = ""
    chkInsurance.Value = 1

    cboCashModeofPayment = ""

    txtVehicleMake = ""
    txtVehicleYear = ""
    txtVehicleColor = ""
    txtCusCode = ""
    txtVehicleVinNo = ""
    txtVehicleDescription = ""
    txtVehicleModel = ""
    txtVehicleModelCode = ""
    cboSalesAE = ""
    cboPurchaseType = ""
    dtDateInvoiced = LOGDATE
    chkInsurance = 0
    cboInsuranceCompany = ""
    txtVehicleSerialNo = ""

    'cboSalesOrderNo.Enabled = False
    txtPreparedBy = ""
    txtCheckedBy = ""
    txtSalesApproved = ""
    txtGeneralManager = ""
    lvAccesories.ListItems.Clear
    txtTotalAccesories = "0.00"
    txtBankPo.Text = ""

    '
    '
    '    Set RSSIGNATORIES = New ADODB.Recordset
    '    RSSIGNATORIES.Open "select * from SMIS_Signatories", gconDMIS, adOpenForwardOnly, adLockReadOnly
    '    If Not RSSIGNATORIES.EOF And Not RSSIGNATORIES.BOF Then
    '        txtPreparedBy = Null2String(RSSIGNATORIES!PreparedBy)
    '        txtCheckedBy = Null2String(RSSIGNATORIES!CheckedBy)
    '        txtSalesApproved = Null2String(RSSIGNATORIES!SalesApproved)
    '        txtGeneralManager = Null2String(RSSIGNATORIES!SalesDispatcher)
    '    End If


End Sub

Sub rsRefresh()
    Set rsInvoice = New ADODB.Recordset
    rsInvoice.Open "select * from SMIS_SalesOrder WHERE VI_NO is not null order by id DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub SearchInvoice(XXX)
    On Error GoTo ErrorCode
    rsInvoice.MoveFirst
    rsInvoice.Find "ID = '" & XXX & "'"
    If (rsInvoice.BOF = True) Or (rsInvoice.EOF = True) Then
        MsgBox "Record not found"
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowCantFind XXX

End Sub

Sub StoreMemVars()
    Dim lngcount                                                      As Integer
    labSJ = "": labORNo = "": labDetails = "": labStatus = "": labInvoiceStatus = ""

    If Not rsInvoice.EOF And Not rsInvoice.BOF Then

        labid.Caption = rsInvoice!ID
        labProspectID = rsInvoice!PROSPECTID

        If Null2String(rsInvoice!ZERORATED) = "" Then
            chkZeroRate2.Value = 0
            chkZeroRate1 = 0
        Else
            If rsInvoice!ZERORATED = False Then
                chkZeroRate2.Value = 0
                chkZeroRate1 = 0
            Else
                chkZeroRate2.Value = 1
                chkZeroRate1 = 1
            End If
        End If

        LABALLOWREPRINT = Null2String(rsInvoice!PRINTEDVI)
        cboSalesOrderNo = Null2String(rsInvoice!SO_NO)
        labVDRNo = Null2String(rsInvoice!VDR_NO)
        txtVINO = Null2String(rsInvoice!VI_NO)
        dtDateInvoiced.Value = Null2String(rsInvoice!InvoicedDate)
        txtVehicleDateReleased = Null2String(rsInvoice!DateReleased)
        cboPurchaseType = Null2String(rsInvoice!Type)
        cboSalesAE = Null2String(rsInvoice!salesae)
        txtSODate = Null2String(rsInvoice!DEYT)
        txtRelease_VDR = Null2String(rsInvoice!VDR_NO)
        txtRelease_Date = Null2String(rsInvoice!DateReleased)

        If rsInvoice!TERM = "COD" Then
            cboPaymentTerm.ListIndex = 0
        ElseIf rsInvoice!TERM = "F" Then
            cboPaymentTerm.ListIndex = 1
        ElseIf rsInvoice!TERM = "BPO" Then
            cboPaymentTerm.ListIndex = 2
        Else
            cboPaymentTerm.ListIndex = 3
        End If

        ''CUSTOMER
        txtCusCode = Null2String(rsInvoice!CODE)
        txtCustName = Null2String(rsInvoice!CustName)
        txtContactPerson = Null2String(rsInvoice!Person)
        txtSpouse = Null2String(rsInvoice!Spouse)
        txtDateBirth = Null2String(rsInvoice!BirthDate)
        txtPosition = Null2String(rsInvoice!posisyon)
        txtHomeAdd = Null2String(rsInvoice!HomeAddress)
        txtTelephoneHome = Null2String(rsInvoice!HomeTelNo)
        txtOfficeAdd = Null2String(rsInvoice!OfficeAdd)
        txtTelephoneOffice = Null2String(rsInvoice!officetelno)
        txtTin = Null2String(rsInvoice!TIN)
        txtCTCNo = Null2String(rsInvoice!CtcNo)
        txtIssuedAt = Null2String(rsInvoice!IssuedAt)
        txtIssuedOn = Null2String(rsInvoice!IssuedOn)
        txtDeliveryAddress = Null2String(rsInvoice!DELIVERY_ADDRESS)
        txtDeliveryInstruction = Null2String(rsInvoice!DELIVERY_INSTRUCTION)
        cboInsuranceCompany = Null2String(rsInvoice!INSURANCECOMPANY)
'updated by: IEBV 0942010_030pm
'-----------------------------------------------------------------------
        If COMPANY_CODE = "HPC" Or COMPANY_CODE = "HSR" Then
                If cboPaymentTerm.Text = "Bank PO" Or cboPaymentTerm.Text = "Company PO" Then
                    txtBankPo.Visible = True
                    Label6.Visible = True
                    Label5.Visible = True
                    dtbankcom_po.Visible = True
                    txtBankPo.Text = Null2String(rsInvoice!BANK_COM_PO)
                  If IsNull(rsInvoice!BANK_COM_PO_DATE) = True Then
                    dtbankcom_po.Value = LOGDATE
                  Else
                    dtbankcom_po.Value = CDate(rsInvoice!BANK_COM_PO_DATE)
                  End If
                End If
        Else
                dtbankcom_po.Visible = False
                Label6.Visible = False
                Label5.Visible = False
                txtBankPo.Text = ""
                txtBankPo.Visible = False
        End If
'-----------------------------------------------------------------------
        If Null2String(rsInvoice!Insured) = "I" Then
            chkInsurance.Value = 1
            If IsDate(rsInvoice("INSUREDDATE")) = True Then
                DTPicker2.Value = rsInvoice("INSUREDDATE")
            Else
                DTPicker2.Value = Date
            End If
        Else
            chkInsurance.Value = 0

        End If
        'CASH TERMS
        cboCashModeofPayment = GetPo(Null2String(rsInvoice!modeofpayment))
        txtCashSalesPrice = FormatNumber(NumericVal(rsInvoice!SALESPRICE))
        txtCashNetSalesPrice = FormatNumber(NumericVal(rsInvoice!NETSALESPRICE))
        txtCashInsurance = FormatNumber(NumericVal(rsInvoice!INSURANCE))
        txtCashLTORegFee = FormatNumber(NumericVal(rsInvoice!LTOREGFEE))
        txtCashAccessories = FormatNumber(NumericVal(rsInvoice!Accessories))
        txtCashDiscount = FormatNumber(NumericVal(rsInvoice!DISCOUNT))
        LAB_TOTAL_CASH = FormatNumber(NumericVal(rsInvoice!Total))
        txtCashFreight = FormatNumber(NumericVal(rsInvoice!FREIGHT))
        txtCashOthersDesc = Null2String(rsInvoice!OTHERSDESC)
        txtCashOthers = FormatNumber(NumericVal(rsInvoice!OTHERS))
        txtCashTax = FormatNumber(NumericVal(rsInvoice!TAX))
        'FINANCING TERMS
        txtFinModeofPayment = Null2String(rsInvoice!modeofpayment)
        txtFinSalesPrice = FormatNumber(NumericVal(rsInvoice!SALESPRICE))
        txtFinNetSalesPrice = FormatNumber(NumericVal(rsInvoice!NETSALESPRICE))
        txtFinDownPayment = FormatNumber(NumericVal(rsInvoice!DownPayment))
        txtFinDownpaymentRate = FormatNumber(NumericVal(rsInvoice!DOWNPAYMENTRATE))
        txtFinInsurance = FormatNumber(NumericVal(rsInvoice!INSURANCE))
        txtFinLTORegFee = FormatNumber(NumericVal(rsInvoice!LTOREGFEE))
        txtFinChattel = FormatNumber(NumericVal(rsInvoice!CHMOFEE))
        txtFinAccessories = FormatNumber(NumericVal(rsInvoice!Accessories))
        txtFinFreight = FormatNumber(NumericVal(rsInvoice!FREIGHT))
        txtFinOthersDesc = Null2String(rsInvoice!OTHERSDESC)
        txtFinOthers = FormatNumber(NumericVal(rsInvoice!OTHERS))
        txtFinBaltoFinanced = FormatNumber(NumericVal(rsInvoice!BALTOFINANCED))
        txtFinDiscount = FormatNumber(NumericVal(rsInvoice!DISCOUNT))
        LAB_TOTAL_FIN = FormatNumber(NumericVal(rsInvoice!Total))
        txtFinGMI = FormatNumber(NumericVal(rsInvoice!GMI))
        txtFinRPPD = FormatNumber(NumericVal(rsInvoice!RPPD))
        txtFinNoOfTermAmort = FormatNumber(NumericVal(rsInvoice!MONTHSAMORT))
        txtFinNetMonthlyAmort = FormatNumber(NumericVal(rsInvoice!NETMOAMORT))
        infoAdditionalInfo = Null2String(rsInvoice!ADDITIONALINFO)
        txtFinAOR = FormatNumber(NumericVal(rsInvoice!AOR), 3)
        txtFinBankTerm = Null2String(rsInvoice!BANKTERM)
        txtFinTax = FormatNumber(NumericVal(rsInvoice!TAX))
        cboFinFinancingCo = Null2String(rsInvoice!financingco)

        txtPreparedBy = Null2String(rsInvoice!PreparedBy)
        txtCheckedBy = Null2String(rsInvoice!CheckedBy)
        txtSalesApproved = Null2String(rsInvoice!SalesApproved)
        txtGeneralManager = Null2String(rsInvoice!SalesDispatcher)
        Label48 = Null2String(rsInvoice!ReasonCancel)
        txtReasonCancel = Null2String(rsInvoice!ReasonCancel)
        cboAccountType = GetAccountType(Null2String(rsInvoice!AccountType))
        'labORNo = CheckORNum(Null2String(rsInvoice!VI_NO), "VI")
        'labSJ = CheckSJNum(Null2String(rsInvoice!VI_NO), "VI")


        If Null2String(rsInvoice!STATUS) <> "C" Then
            If Null2String(rsInvoice!DateReleased) <> "" Then

                labInvoiceStatus = "**RELEASED**"

                labORNo = CheckORNum(Null2String(rsInvoice!VI_NO), "VI")
                labSJ = CheckSJNum(Null2String(rsInvoice!VI_NO), "VI")

                If Null2String(rsInvoice!STATUS) = "P" Then
                    labStatus = "***POSTED ***"
                    cmdCancelCO.Enabled = False
                    cmdUnPost.Enabled = True
                    cmdPost.Enabled = False
                    cmdPrint.Enabled = True
                    cmdEdit.Enabled = False
                    cmdRelease.Enabled = False
                    cmdUnReleased.Enabled = True
                    If labORNo <> "" And labSJ <> "" Then
                        labDetails = " OR ISSUED/IMPORTED TRANSACTION"
                        cmdUnPost.Enabled = False
                        cmdPrint.Enabled = False
                        cmdCancelCO.Enabled = False
                        cmdUnReleased.Enabled = False
                    ElseIf labORNo = "" And labSJ <> "" Then
                        labDetails = "IMPORTED TRANSACTION"
                        cmdUnPost.Enabled = False
                        cmdPrint.Enabled = False
                        cmdCancelCO.Enabled = False
                        cmdUnReleased.Enabled = False
                    ElseIf labORNo <> "" And labSJ = "" Then
                        labDetails = "OR ISSUED"
                        cmdUnPost.Enabled = False
                        cmdPrint.Enabled = False
                        cmdCancelCO.Enabled = False
                        cmdUnReleased.Enabled = False
                    End If
                Else
                    cmdCancelCO.Enabled = True
                    cmdUnPost.Enabled = False
                    cmdPost.Enabled = True
                    cmdEdit.Enabled = True
                    cmdRelease.Enabled = False
                    cmdUnReleased.Enabled = True
                    cmdPrint.Enabled = True
                End If

                fraHeader.Enabled = False
                fraTermsCash.Enabled = False
                fraTermsCredit.Enabled = False
                fraPrintingDetails.Enabled = True
                fraAccessories.Enabled = False
                cmdEditCustInfo.Enabled = False
                cmdSelectVehicles.Enabled = False
                cboSalesOrderNo.Enabled = False
            Else
                labInvoiceStatus = "**ON PROCESS**"
                cmdCancelCO.Enabled = True
                cmdUnPost.Enabled = False
                cmdPost.Enabled = True
                cmdEdit.Enabled = True
                cmdRelease.Enabled = True
                cmdUnReleased.Enabled = False
                fraTermsCash.Enabled = True
                fraTermsCredit.Enabled = True
                fraHeader.Enabled = True
                fraPrintingDetails.Enabled = True
                fraAccessories.Enabled = True
                cmdSelectVehicles.Enabled = True
                cmdEditCustInfo.Enabled = True
                cboSalesAE.Enabled = True
                cboSalesOrderNo.Enabled = True
            End If
        Else
            cmdCancelCO.Enabled = False:
            cmdUnPost.Enabled = False
            cmdPost.Enabled = False
            labStatus = "***CANCELLED***"
            cmdRelease.Enabled = False
            cmdUnReleased.Enabled = False
            cmdEdit.Enabled = False
            cmdPrint.Enabled = False
            labInvoiceStatus = ""
            labInvoiceStatus = ""
        End If



        Dim rsMrr                                                     As ADODB.Recordset
        Set rsMrr = gconDMIS.Execute("SELECT * FROM SMIS_MRRINV_TABLE WHERE IGNKEY=" & N2Str2Null(rsInvoice!IGNKEY_NO))
        lblVehicleInformation = vbNullString
        txtOldCS = Null2String(rsInvoice!IGNKEY_NO)
        If Not (rsMrr.EOF Or rsMrr.BOF) Then
            txtMCode = rsMrr.Fields("ID")
            txtVehicleMake = Null2String(rsMrr!Make)
            txtVehicleModel = Null2String(rsMrr!Model)
            txtVehicleDescription = Null2String(rsMrr!DESCRIPT)
            txtVehicleYear = Null2String(rsMrr!YEER)
            txtVehicleConductionSticker = Null2String(rsMrr!ignkey)

            txtVehicleEngineNo = Null2String(rsMrr!EngineNo)
            txtVehicleFrameNo = Null2String(rsMrr!frameno)
            txtVehicleProdNo = Null2String(rsMrr!prodno)
            txtVehicleColor = Null2String(rsMrr!Color)
            txtVehicleVinNo = Null2String(rsMrr!VINO)
            txtVehicleSerialNo = Null2String(rsMrr!SERIALNO)

            lblVehicleInformation = Null2String(rsMrr!DESCRIPT)
            txtVehicleTransmission = Null2String(rsMrr!Transmission)
            txtVehicleModelCode = Null2String(rsMrr!ModelCode)
            'ACCESSORIES INFO


            If txtVehicleConductionSticker <> "" Then
                Dim totalacc                                          As Double
                Dim temprs                                            As ADODB.Recordset
                'UPDATED BY: JUN
                'DATE UPDATED: 08/05/2008
                If COMPANY_CODE = "HAS" Then
                    Set temprs = gconDMIS.Execute("Select Description ,QTY , COST , QTY * COST, ID,ISFREE from SMIS_MRRINV_DETAIL Where IgnKeyNo =" & N2Str2Null(txtVehicleConductionSticker))
                Else
                    Set temprs = gconDMIS.Execute("Select Description ,QTY , COST , QTY * COST, ID from SMIS_MRRINV_DETAIL Where IgnKeyNo =" & N2Str2Null(txtVehicleConductionSticker))
                End If

                Dim lst                                               As ListItem
                lvAccesories.ListItems.Clear
                'UPDATED BY: JUN
                'DATE UPDATED: 08/05/2008
                If COMPANY_CODE = "HAS" Then
                    While Not temprs.EOF
                        Set lst = lvAccesories.ListItems.Add(, , Null2String(temprs.Fields(0).Value))
                        lst.ListSubItems.Add , , NumericVal(temprs.Fields(1).Value)
                        lst.ListSubItems.Add , , FormatNumber(NumericVal(temprs.Fields(2).Value))
                        lst.ListSubItems.Add , , FormatNumber(NumericVal(temprs.Fields(3).Value))
                        totalacc = totalacc + (NumericVal(temprs.Fields(3).Value))
                        lst.ListSubItems.Add , , temprs.Fields(4).Value
                        lst.ListSubItems.Add , , Abs(temprs.Fields(5).Value)
                        temprs.MoveNext
                    Wend
                Else
                    While Not temprs.EOF
                        Set lst = lvAccesories.ListItems.Add(, , Null2String(temprs.Fields(0).Value))
                        lst.ListSubItems.Add , , NumericVal(temprs.Fields(1).Value)
                        lst.ListSubItems.Add , , FormatNumber(NumericVal(temprs.Fields(2).Value))
                        lst.ListSubItems.Add , , FormatNumber(NumericVal(temprs.Fields(3).Value))
                        totalacc = totalacc + (NumericVal(temprs.Fields(3).Value))
                        lst.ListSubItems.Add , , temprs.Fields(4).Value
                        temprs.MoveNext
                    Wend
                End If

                Set temprs = Nothing
                Set lst = Nothing
            Else
                lvAccesories.ListItems.Clear
            End If
            txtFinAccessories = totalacc
            txtCashAccessories = totalacc
            txtTotalAccesories = totalacc

        Else
            txtVehicleConductionSticker = Null2String(rsInvoice!IGNKEY_NO)
            txtVehicleModel = Null2String(rsInvoice!Model)
            txtVehicleDescription = Null2String(rsInvoice!modeldescription)
            txtVehicleEngineNo = Null2String(rsInvoice!EngineNo)
            txtVehicleFrameNo = Null2String(rsInvoice!frameno)
            txtVehicleVinNo = Null2String(rsInvoice!VINO)
            txtVehicleProdNo = Null2String(rsInvoice!prodno)    '
            txtVehicleColor = Null2String(rsInvoice!Color)
            txtMCode = 0
        End If
        txtVehicleWarrantyCertifcate = Null2String(rsInvoice!certific8)
        txtVehiclePlateNo = Null2String(rsInvoice!PLATE_NO)
    Else
        lngcount = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_SALESORDER WHERE VI_NO IS  NULL AND STATUS<>'C' OR STATUS IS NULL").Fields(0).Value
        If lngcount = 0 Then
            Call MsgBox("There are no Sales Order to Process for Invoicing.. " _
                      & vbCrLf & "Please Add Sales Order to process invoicing. " _
                      & vbCrLf & "Vehicle Invoice Will Now Unload" _
               , vbInformation, "No Sales Orders")
            Unload Me
            Exit Sub
        Else
            ShowNoRecord
            If MsgBox("Do You Want to Add New Vehcile Invoice", vbOKCancel + vbQuestion) = vbOK Then
                cmdAdd.Value = True
            Else
                Unload Me
            End If
        End If
    End If
End Sub

Sub UpdateAccessoriesAmount()
    txtAccAmount = FormatNumber(NumericVal(txtAccQty) * NumericVal(txtAccRate))
End Sub

Sub UpdateTotalAmount()
    If AddorEdit = "" Then: Exit Sub
    Dim A, b, C, D, E, F, G, H, i, j
    If COMPANY_CODE = "HNC" Then
        A = NumericVal(txtFinSalesPrice)
    Else
        A = NumericVal(txtFinNetSalesPrice)
    End If
    b = NumericVal(txtFinDownPayment)


    D = NumericVal(txtFinInsurance)
    E = NumericVal(txtFinLTORegFee)
    F = NumericVal(txtFinFreight)
    G = NumericVal(txtFinOthers)

    i = NumericVal(txtFinChattel)
    j = NumericVal(txtFinDownpaymentRate)
    C = NumericVal(txtFinDiscount)

    If UCase(cboPaymentTerm) = "FINANCING" Or UCase(cboPaymentTerm) = "BANK PO" Then
        Tutal = A - (b)
    Else
        Tutal = A - (C + b)
    End If

    txtFinBaltoFinanced = FormatNumber(Tutal)
    txtFinNetSalesPrice = FormatNumber(NumericVal(txtFinSalesPrice) - NumericVal(txtFinDiscount))
    LAB_TOTAL_FIN = FormatNumber((b + D + E + F + G + i))

End Sub

Sub UpdateTotalCashAmount()
    If AddorEdit = "" Then Exit Sub
    On Error Resume Next
    Dim A, b, C, D, E, F, G, H
    A = NumericVal(txtCashSalesPrice)
    b = NumericVal(txtCashDiscount)
    C = NumericVal(txtCashInsurance)
    D = NumericVal(txtCashLTORegFee)
    E = NumericVal(txtCashFreight)
    F = NumericVal(txtCashOthers)
    LAB_TOTAL_CASH = FormatNumber(A + C + D + E + F + H - (b), 2)
    txtCashNetSalesPrice = FormatNumber(A - b)
End Sub

Sub z_UpdateAccountCode()
    Dim SQL                                                           As String
    Dim rsA                                                           As ADODB.Recordset
    If Not txtVehicleModel = "" Then

        Set rsA = gconDMIS.Execute("Select TOP 1  ATC_SALESDISC_FLEET, ATC_SALESDISC_RETAIL ,ATC_SALES_FLEET ,ATC_SALES_RETAIL ,ATC_COSTOFSALES_FLEET ,ATC_COSTOFSALES_RETAIL ,ATC_INVENTORY  from ALL_MODEL WHERE MODEL='" & txtVehicleModel & "'")
        If Not rsA.BOF Or Not rsA.BOF Then

            SQL = " UPDATE SMIS_SALESORDER SET "
            SQL = SQL & " ATC_SALESDISC_FLEET=" & N2Str2Null(Null2String(rsA!ATC_SALESDISC_FLEET)) & ","
            SQL = SQL & " ATC_SALESDISC_RETAIL =" & N2Str2Null(Null2String(rsA!ATC_SALESDISC_RETAIL)) & ","

            SQL = SQL & " ATC_SALES_FLEET=" & N2Str2Null(Null2String(rsA!ATC_SALES_FLEET)) & ","
            SQL = SQL & " ATC_SALES_RETAIL=" & N2Str2Null(Null2String(rsA!ATC_SALES_RETAIL)) & ","

            SQL = SQL & " ATC_COSTOFSALES_FLEET=" & N2Str2Null(Null2String(rsA!ATC_COSTOFSALES_FLEET)) & ","
            SQL = SQL & " ATC_COSTOFSALES_RETAIL=" & N2Str2Null(Null2String(rsA!ATC_COSTOFSALES_RETAIL)) & ","

            SQL = SQL & " ATC_INVENTORY=" & N2Str2Null(Null2String(rsA!ATC_INVENTORY))
            SQL = SQL & " WHERE VI_NO='" & txtVINO & "'"
            gconDMIS.Execute SQL
        End If
    End If
End Sub

Sub z_UpdatePDICheckList()
    'UDPATING CODE      :   AXP-065082007 328PM
    Dim vtxtVINo                                                      As String
    Dim vtxtCustName                                                  As String
    Dim vtxtVehiclePlateNo                                            As String
    Dim vtxtVehicleMake                                               As String
    Dim vtxtVehicleModel                                              As String
    Dim vtxtVehicleDescription                                        As String
    Dim vtxtVehicleEngineNo                                           As String
    Dim vtxtVehicleVinNo                                              As String
    Dim vtxtVehicleColor                                              As String
    Dim vtxtVehicleTransmission                                       As String
    Dim vcboSalesAE                                                   As String
    Dim vtxtVehicleModelcode
    vtxtVINo = N2Str2Null(txtVINO)
    vtxtCustName = N2Str2Null(txtCustName)
    vtxtVehiclePlateNo = N2Str2Null(txtVehiclePlateNo)
    vtxtVehicleMake = N2Str2Null(txtVehicleMake)
    vtxtVehicleModel = N2Str2Null(txtVehicleModel)
    vtxtVehicleDescription = N2Str2Null(txtVehicleDescription)
    vtxtVehicleEngineNo = N2Str2Null(txtVehicleEngineNo)
    vtxtVehicleVinNo = N2Str2Null(txtVehicleVinNo)
    vtxtVehicleColor = N2Str2Null(txtVehicleColor)
    vtxtVehicleTransmission = N2Str2Null(txtVehicleTransmission)
    vcboSalesAE = N2Str2Null(cboSalesAE)
    vtxtVehicleDescription = N2Str2Null(txtVehicleDescription)
    vtxtVehicleModelcode = N2Str2Null(txtVehicleModelCode)
    Dim SQL                                                           As String

    If gconDMIS.Execute("Select COUNT(*) FROM SMIS_PDI_HDR WHERE VI_NO=" & vtxtVINo).Fields(0).Value = 0 Then
        GoTo ER1
    End If

    If AddorEdit = "ADD" Then

ER1:         SQL = " INSERT INTO SMIS_PDI_HDR ("
        SQL = SQL & " PDIDate , VI_NO , "
        SQL = SQL & " CustName, PlateNo, Make,Model, ModelCode,"
        SQL = SQL & " ModelDescription,"
        SQL = SQL & " EngineNo, Vino, Color, "
        SQL = SQL & " Tranmission,  "
        SQL = SQL & " SAE) VALUES("
        SQL = SQL & N2Str2Null(LOGDATE) & ","
        SQL = SQL & vtxtVINo & ","
        SQL = SQL & vtxtCustName & ","
        SQL = SQL & vtxtVehiclePlateNo & ","
        SQL = SQL & vtxtVehicleMake & ","
        SQL = SQL & vtxtVehicleModel & ","
        SQL = SQL & vtxtVehicleModelcode & ","
        SQL = SQL & vtxtVehicleDescription & ","
        SQL = SQL & vtxtVehicleEngineNo & ","
        SQL = SQL & vtxtVehicleVinNo & ","
        SQL = SQL & vtxtVehicleColor & ","
        SQL = SQL & vtxtVehicleTransmission & ","
        SQL = SQL & vcboSalesAE & ")"
        gconDMIS.Execute (SQL)

        SQL = " INSERT INTO SMIS_PDI_DET " & vbCrLf
        SQL = SQL & "select " & vtxtVINo & " , PDILINEID, 0,'N' AS STATUS from SMIS_vw_PDILookUp where modelcode=" & vtxtVehicleModelcode
        gconDMIS.Execute (SQL)
        LogAudit "A", "PDI CHECKLIST", "CUSTOMER NAME " & txtCustName & " MODEL & txtVehicleModel " & " VIN" & txtVehicleVinNo
    Else
        SQL = " Update SMIS_PDI_HDR SET "
        SQL = SQL & "VI_NO=" & vtxtVINo & " ,"
        SQL = SQL & "CustName=" & vtxtCustName & " ,"
        SQL = SQL & "PlateNo=" & vtxtVehiclePlateNo & " ,"
        SQL = SQL & "Make=" & vtxtVehicleMake & " ,"
        SQL = SQL & "Model=" & vtxtVehicleModel & " ,"
        SQL = SQL & "ModelCode=" & vtxtVehicleModelcode & " ,"
        SQL = SQL & "ModelDescription=" & vtxtVehicleDescription & " ,"
        SQL = SQL & "EngineNo=" & vtxtVehicleEngineNo & " ,"
        SQL = SQL & "Vino=" & vtxtVehicleVinNo & " ,"
        SQL = SQL & "Color=" & vtxtVehicleColor & " ,"
        SQL = SQL & "Tranmission=" & vtxtVehicleTransmission & " ,"
        SQL = SQL & "SAE=" & vcboSalesAE
        SQL = SQL & "WHERE VI_NO= " & vtxtVINo
        gconDMIS.Execute (SQL)
        LogAudit "E", "PDI CHECKLIST", "CUSTOMER NAME " & txtCustName & " MODEL & txtVehicleModel " & " VIN" & txtVehicleVinNo
    End If
End Sub

Private Sub cboAccessories_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboFinFinancingCo_Change()
    'If AddorEdit = "" Or AddorEdit = "ADD" Then: Exit Sub
    'txtFinNetMonthlyAmort = AORVALUE(NumericVal(txtFinBaltoFinanced), NumericVal(txtFinAOR), NumericVal(txtFinNoOfTermAmort))
End Sub

Private Sub cboFinFinancingCo_Click()
    cboFinFinancingCo_Change
End Sub

Private Sub cboPaymentTerm_Change()
    cboPaymentTerm_Click
End Sub

Private Sub cboPaymentTerm_Click()

    If cboPaymentTerm.ListIndex = 0 Or cboPaymentTerm.ListIndex = 3 Then
        cboFinFinancingCo.Enabled = False
        txtFinBankTerm.Enabled = False
        fraTermsCash.Visible = True
        fraTermsCredit.Visible = False
        SSTabVDetails.Item(2).Caption = "Cash Terms"
        UpdateTotalCashAmount
    Else
        fraTermsCash.Visible = False
        fraTermsCredit.Visible = True
        cboFinFinancingCo.Enabled = True
        SSTabVDetails.Item(2).Caption = "Financing Terms"
        txtFinBankTerm.Enabled = True
        UpdateTotalAmount
    End If
    If COMPANY_CODE = "HPC" Or COMPANY_CODE = "HSR" Then
        If cboPaymentTerm.Text = "Bank PO" Or cboPaymentTerm.Text = "Company PO" Then
            txtBankPo.Visible = True
            dtbankcom_po.Visible = True
            Label6.Visible = True
            Label5.Visible = True
        Else
            txtBankPo.Visible = False
            dtbankcom_po.Visible = False
            Label6.Visible = False
            Label5.Visible = False
        End If
    End If
End Sub

Private Sub cboSalesAE_LostFocus()
    cboSalesAE.ListIndex = SelectCombo(cboSalesAE, cboSalesAE, False)
    If cboSalesAE.ListIndex = -1 Then: cboSalesAE = ""
End Sub

Private Sub cboSalesOrderNo_Change()
    If cboSalesOrderNo.ListIndex = -1 Or ((AddorEdit = "EDIT" Or AddorEdit = "")) Then: Exit Sub
    Dim temprs                                                        As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT * FROM SMIS_SALESORDER WHERE ID=" & cboSalesOrderNo.ItemData(cboSalesOrderNo.ListIndex))

    If Not temprs.EOF And Not temprs.BOF Then

        labid.Caption = temprs!ID

        labProspectID = Null2String(temprs!PROSPECTID)
        txtVehicleDateReleased = Null2String(temprs!DateReleased)
        txtCusCode = Null2String(temprs!CODE)
        dtDateInvoiced.Value = Now
        txtSODate = Null2String(temprs!DEYT)
        cboSalesOrderNo = Null2String(temprs!SO_NO)
        cboPurchaseType = Null2String(temprs!Type)
        txtCustName = Null2String(temprs!CustName)
        txtHomeAdd = Null2String(temprs!HomeAddress)
        txtTelephoneHome = Null2String(temprs!HomeTelNo)
        txtOfficeAdd = Null2String(temprs!OfficeAdd)
        txtTelephoneOffice = Null2String(temprs!officetelno)
        txtDateBirth = Null2String(temprs!BirthDate)
        txtSpouse = Null2String(temprs!Spouse)
        txtContactPerson = Null2String(temprs!Person)
        txtPosition = Null2String(temprs!posisyon)
        txtTin = Null2String(temprs!TIN)
        txtCTCNo = Null2String(temprs!CtcNo)
        txtIssuedAt = Null2String(temprs!IssuedAt)
        txtIssuedOn = Null2String(temprs!IssuedOn)
        If temprs!TERM = "COD" Then
            cboPaymentTerm.ListIndex = 0
        ElseIf temprs!TERM = "BPO" Then
            cboPaymentTerm.ListIndex = 2
        Else
            cboPaymentTerm.ListIndex = 1
        End If
        lblVehicleInformation = Null2String(temprs!modeldescription)
        txtFinBankTerm = Null2String(temprs!BANKTERM)
        cboFinFinancingCo = Null2String(temprs!financingco)
        cboSalesAE = Null2String(temprs!salesae)
        txtCashSalesPrice = FormatNumber(NumericVal(temprs!SALESPRICE))
        txtCashNetSalesPrice = FormatNumber(NumericVal(temprs!NETSALESPRICE))
        txtCashInsurance = FormatNumber(NumericVal(temprs!INSURANCE))
        txtCashLTORegFee = FormatNumber(NumericVal(temprs!LTOREGFEE))
        txtCashAccessories = FormatNumber(NumericVal(temprs!Accessories))
        cboCashModeofPayment = FormatNumber(NumericVal(temprs!modeofpayment))
        txtCashFreight = FormatNumber(NumericVal(temprs!FREIGHT))
        txtCashOthersDesc = Null2String(temprs!OTHERSDESC)
        txtCashOthers = FormatNumber(NumericVal(temprs!OTHERS))
        txtCashDiscount = FormatNumber(NumericVal(temprs!DISCOUNT))
        txtFinSalesPrice = FormatNumber(NumericVal(temprs!SALESPRICE))
        txtFinNetSalesPrice = FormatNumber(NumericVal(temprs!NETSALESPRICE))
        txtFinDownPayment = FormatNumber(NumericVal(temprs!DownPayment))
        txtFinDiscount = FormatNumber(NumericVal(temprs!DISCOUNT))
        txtFinInsurance = FormatNumber(NumericVal(temprs!INSURANCE))
        txtFinLTORegFee = FormatNumber(NumericVal(temprs!LTOREGFEE))
        txtFinChattel = FormatNumber(NumericVal((temprs!CHMOFEE)))
        txtFinAccessories = FormatNumber(NumericVal(temprs!Accessories))
        txtFinFreight = FormatNumber(NumericVal(temprs!FREIGHT))
        txtFinOthersDesc = Null2String(temprs!OTHERSDESC)
        txtFinOthers = FormatNumber(NumericVal(temprs!OTHERS))
        txtFinBaltoFinanced = FormatNumber(NumericVal(temprs!BALTOFINANCED))
        txtFinGMI = FormatNumber(NumericVal(temprs!GMI))
        txtFinRPPD = FormatNumber(NumericVal(temprs!RPPD))
        txtFinNoOfTermAmort = FormatNumber(NumericVal(temprs!MONTHSAMORT))
        txtFinNetMonthlyAmort = FormatNumber(NumericVal(temprs!NETMOAMORT))
        LAB_TOTAL_FIN = FormatNumber(NumericVal(temprs!Total))
        txtFinBankTerm = FormatNumber(NumericVal(temprs!MONTHSAMORT))
        txtFinAOR = FormatNumber(NumericVal(temprs!AOR))
        txtVehicleConductionSticker = Null2String(temprs!IGNKEY_NO)
        txtVehicleModel = Null2String(temprs!Model)
        lblVehicleInformation = Null2String(temprs!modeldescription)
        txtVehicleDescription = Null2String(temprs!modeldescription)
        txtVehicleProdNo = Null2String(temprs!prodno)
        txtVehicleEngineNo = Null2String(temprs!EngineNo)
        txtVehicleFrameNo = Null2String(temprs!frameno)
        txtVehicleVinNo = Null2String(temprs!VINO)
        txtVehicleColor = Null2String(temprs!Color)
        txtVehiclePlateNo = Null2String(temprs!PLATE_NO)
        txtOldCS = Null2String(temprs!IGNKEY_NO)
        If temprs!Insured = "I" Then
            chkInsurance.Value = 1
        Else
            chkInsurance.Value = 0
        End If
        txtMCode = 0
        If txtVehicleConductionSticker <> "" Then
            Dim rssMRRID                                              As ADODB.Recordset
            Set rssMRRID = gconDMIS.Execute("SELECT * FROM SMIS_MRRINV_TABLE WHERE IGNKEY='" & txtVehicleConductionSticker & "'")
            If Not (rssMRRID.EOF Or rssMRRID.BOF) Then
                lblVehicleStatus = " Available"
                txtMCode = Null2String(temprs!ID)
                txtVehicleYear = Null2String(rssMRRID!YEER)
                txtVehicleMake = Null2String(rssMRRID!Make)
                txtVehicleSerialNo = Null2String(rssMRRID!SERIALNO)
                txtVehicleTransmission = Null2String(rssMRRID!Transmission)
                txtVehicleModelCode = GetModelCode(txtVehicleModel)
            Else
                lblVehicleStatus = " Insufficient Vehicles Informations .. Please Update"
                txtVehicleYear = ""
                txtMCode = 0
                txtMCode = ""
                txtVehicleYear = ""
                txtVehicleMake = ""
                txtVehicleSerialNo = ""
            End If
        Else
            lblVehicleStatus = " Insufficient Vehicles Informations .. Please Update"
        End If
        infoAdditionalInfo = Null2String(temprs!ADDITIONALINFO)
        Set rsSignatories = New ADODB.Recordset
        rsSignatories.Open "select * from SMIS_Signatories where usedin='SALES INVOICE'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsSignatories.EOF And Not rsSignatories.BOF Then
            txtGeneralManager = Null2String(rsSignatories!GeneralManager)
            txtPreparedBy = Null2String(rsSignatories!PreparedBy)
            txtCheckedBy = Null2String(rsSignatories!CheckedBy)
            txtSalesApproved = Null2String(rsSignatories!SalesApproved)
        End If
        If txtVehicleConductionSticker <> "" Then
            Dim totalacc                                              As Double
            Dim rsacc                                                 As ADODB.Recordset
            Set rsacc = gconDMIS.Execute("Select Description ,QTY , COST , QTY * COST, ID from SMIS_MRRINV_DETAIL Where IgnKeyNo =" & N2Str2Null(txtVehicleConductionSticker))
            Dim lst                                                   As ListItem
            lvAccesories.ListItems.Clear
            While Not rsacc.EOF
                Set lst = lvAccesories.ListItems.Add(, , Null2String(rsacc.Fields(0).Value))
                lst.ListSubItems.Add , , NumericVal(rsacc.Fields(1).Value)
                lst.ListSubItems.Add , , FormatNumber(NumericVal(rsacc.Fields(2).Value))
                lst.ListSubItems.Add , , FormatNumber(NumericVal(rsacc.Fields(3).Value))
                totalacc = totalacc + (NumericVal(rsacc.Fields(3).Value))
                lst.ListSubItems.Add , , IsNull(rsacc.Fields(4).Value)
                rsacc.MoveNext
            Wend

            Set rsacc = Nothing
            Set lst = Nothing
        Else
            lvAccesories.ListItems.Clear
        End If
        txtFinAccessories = totalacc
        txtCashAccessories = totalacc
        txtTotalAccesories = totalacc
        labStatus = ""
        labInvoiceStatus = ""
        txtRelease_Date = ""
        txtRelease_VDR = ""
        cmdSelectVehicles.Enabled = True
        picHeader.Enabled = True
        picCustomerInformation.Enabled = True
        picTinInfo.Enabled = True
        picPrintingDetails.Enabled = True
        picTerms.Enabled = True
        picViewAccessories.Enabled = True
        picVehiclesDetail.Enabled = True
        cmdEditCustInfo.Enabled = True
        fraTermsCash.Enabled = True
        fraTermsCredit.Enabled = True
        fraAccessories.Enabled = True
        fraPlateno.Enabled = True
        fraPrintingDetails.Enabled = True
    Else
        MsgBox " Cannot Find The Record"
    End If
End Sub

Private Sub cboSalesOrderNo_Click()
    cboSalesOrderNo_Change
End Sub

Private Sub Check1_Click()

    lvViewVehicles.Columns(0).Visible = CBool(Check1.Value)
    lvViewVehicles.Columns(1).Visible = CBool(Check1.Value)
    lvViewVehicles.Columns(5).Visible = CBool(Check1.Value)
    lvViewVehicles.Columns(6).Visible = CBool(Check1.Value)
    lvViewVehicles.Columns(7).Visible = CBool(Check1.Value)
    lvViewVehicles.Columns(9).Visible = CBool(Check1.Value)
End Sub

Private Sub chkInsurance_Click()
    If chkInsurance.Value = 1 Then
        picInsurance.Visible = True
    Else
        picInsurance.Visible = False
    End If
End Sub

Private Sub chkZeroRate1_Click()
    On Error Resume Next
    If chkZeroRate1.Value = 1 Then
        txtCashTax = "0.00"
        txtCashTax.Enabled = False
    Else
        txtCashTax.Enabled = True
        If Not rsInvoice.EOF Or Not rsInvoice.BOF Then
            txtCashTax = FormatNumber(NumericVal(rsInvoice("TAX")))
        End If
    End If
End Sub

Private Sub chkZeroRate2_Click()
    On Error Resume Next
    If chkZeroRate2.Value = 1 Then
        txtFinTax = "0.00"
        txtFinTax.Enabled = False
    Else
        txtFinTax.Enabled = True
        If Not rsInvoice.EOF Or Not rsInvoice.BOF Then
            txtFinTax = FormatNumber(NumericVal(rsInvoice("TAX")))
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "VEHICLES INVOICING") = False Then Exit Sub
    Dim lngcount                                                      As Integer
    On Error GoTo ErrorCode:
    labInvoiceStatus = ""
    labStatus = ""
    lngcount = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_SALESORDER WHERE VI_NO is null and SOSTATUS='P' and  SOSTATUS <>'C'").Fields(0).Value
    If lngcount = 0 Then
        Call MsgBox("There are no Sales Order to Process for Invoicing.. " _
                  & vbCrLf & "Please Add Sales Order to Process invoicing. " _
                  & vbCrLf & "  " _
           , vbInformation, "No Sales Orders")
        Exit Sub
    End If

    AddorEdit = "ADD"
    initMemvars
    picAdds.Visible = False
    picSaves.Visible = True

    cboSalesOrderNo.Enabled = True
    FillCombo "SELECT DISTINCT ID, SO_NO from SMIS_SALESORDER WHERE VI_NO is null  and VDR_NO is NUll and (SOSTATUS='P' and  SOSTATUS <>'C'  ) Order By ID DESC", 0, 1, cboSalesOrderNo
    If cboSalesOrderNo.ListCount > 0 Then
        fraHeader.Enabled = True
        picHeader.Enabled = True
        txtVINO = (GenerateCode("SMIS_SALESORDER", "VI_NO", "000000"))
    Else
        MsgBox " There are no Sales Order To Process", vbInformation
        fraHeader.Enabled = False
        picHeader.Enabled = False
        cboSalesOrderNo.Enabled = False
        cmdCancel.Value = True
    End If
    dtDateInvoiced.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdAddAcc_Click()
    If txtVehicleConductionSticker = "" Then
        MsgBox " Insufficient Vehicle Inventory Information, Please Add Vehicle Inventory Information", vbExclamation
        Exit Sub
    End If
    'Combo_Loadval cboAccessories, gconDMIS.Execute("select distinct upper(Description) from  SMIS_MRRINV_DETAIL where description is not null")
    
    Combo_Loadval cboAccessories, gconDMIS.Execute("select distinct (Description) from  SMIS_MRRINV_DETAIL where description is not null order by description asc")
    txtAccQty = "1"
    txtAccRate = "0.00"
    cboAccessories = ""
    Command5.Enabled = False

    'UPDATED BY: JUN
    'DATE UPDATED: 08/05/2008
    If COMPANY_CODE <> "HAS" Then
        chISFREE.Visible = False
    Else
        chISFREE.Visible = True
    End If
    ShowHidePictureBox2 picAccessories, True
End Sub

Private Sub cmdAuto_Click()
    On Error Resume Next
    If cmdAuto.Value = 1 Then
        txtFinNetMonthlyAmort.Enabled = False
        txtFinNetMonthlyAmort = AORVALUE(NumericVal(txtFinBaltoFinanced), NumericVal(txtFinAOR), NumericVal(txtFinNoOfTermAmort))
        txtFinNetMonthlyAmort = FormatNumber(txtFinNetMonthlyAmort, 2)
    Else
        txtFinNetMonthlyAmort.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    picHeader.Enabled = False
    picCustomerInformation.Enabled = False
    picTinInfo.Enabled = False
    picPrintingDetails.Enabled = False
    picTerms.Enabled = False
    picVehiclesDetail.Enabled = False
    picViewAccessories.Enabled = False
    cmdSelectVehicles.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    cmdAuto.Value = 0
    AddorEdit = ""
    StoreMemVars
    picMultipleSO.Visible = False
End Sub

Private Sub cmdCancelCO_Click()
    On Error GoTo ErrorCode:
    If Function_Access(LOGID, "Acess_CancelEntry", "VEHICLES INVOICING") = False Then Exit Sub

    ShowHidePictureBox2 picCancelReason, True
    On Error Resume Next
    txtReasonCancel.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdCancelDetailProduct_Click(Index As Integer)
    labEDITDetail = "FALSE"
    ShowHidePictureBox2 picAccessories, False

End Sub

Private Sub cmdCancelFinal_Click()
    If LTrim(RTrim(txtReasonCancel)) = "" Then

        MsgSpeechBox "Please input reason for Cancellation of this invoice."

        On Error Resume Next
        txtReasonCancel.SetFocus
        Exit Sub
    End If
    If MsgBox("Do you Want to Cancel this Invoice ", vbOKCancel + vbInformation, "Confirm Posting") = vbCancel Then Exit Sub

    '*********NEW LOG AUDIT
    SQL_STATEMENT = "UPDATE SMIS_SALESORDER SET D_S=GETDATE(),  DATERELEASED = NULL, STATUS='C' , REASONCANCEL=" & N2Str2Null(txtReasonCancel) & "  WHERE ID = " & labid.Caption
    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "C", "VEHICLE INVOICING", SQL_STATEMENT, Null2String(labid), "", "INVOICE No:" & txtVINO, "", ""
    '**********************
    gconDMIS.Execute ("UPDATE SMIS_PDI_HDR SET STATUS='O' WHERE VI_NO=" & N2Str2Null(rsInvoice!IGNKEY_NO))
    
    SQL_STATEMENT = "DELETE FROM CSMS_CUSVEH WHERE CUSCDE=" & N2Str2Null(txtCusCode) & " AND VCOND_NO=" & N2Str2Null(rsInvoice!IGNKEY_NO)
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT ---------------------------------------------------------------
        Call NEW_LogAudit("X", "CUSTOMER VEHICLE", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtCusCode), "CUSCDE", "CSMS_CUSVEH", "DETAILS", N2Str2Null(rsInvoice!IGNKEY_NO), "VCOND_NO"), "", "COND. NO: " & Null2String(rsInvoice!IGNKEY_NO), "", "")
    'NEW LOG AUDIT ---------------------------------------------------------------
    
    Set rsMRRINV = New ADODB.Recordset
    rsMRRINV.Open "Select * from SMIS_MrrInv_Table WHERE ignkey = " & N2Str2Null(rsInvoice!IGNKEY_NO), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
        SQL_STATEMENT = "UPDATE SMIS_MRRINV_TABLE SET RELEASED=0,DATERELEASED = NULL , INVOICEDDATE=null,IStatus='A' WHERE ignkey = " & N2Str2Null(rsInvoice!IGNKEY_NO)
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-------------------------------------------
            'CALL NEW_LogAudit("E",
        'NEW LOG AUDIT-------------------------------------------
        
        MessagePop InfoVoid, "CANCELLED", "RECORD SUCESSFULLY CANCELLED", 1000, 2
        'LogAudit "C", "VEHICLE INVOICE", "NAME" & txtCustName & " VEHICLE:" & txtVehicleVinNo
    Else
        MsgSpeechBox "Warning: MRR Inventory Not Updated! " & vbCrLf & _
                     "Please Manually Update MRR Entry"
    End If



    If rsInvoice("PROSPECTID") > 0 Then
        gconDMIS.Execute ("UPDate CRIS_PROSPECTS  Set InvoiceNo=NULL, LOGCLOSINGDATE= NULL Where PROSPECTID=" & (labProspectID))
        Select Case MsgBox("There is Prospect Information Assosicated with this Sales Invoice." _
                         & vbCrLf & "Do You Want to Set Prospect  Status Active " _
                         & vbCrLf & "Click Yes Make Prospect Active" _
                         & vbCrLf & "Click No To Make Prospect Inactive" _
             , vbYesNo Or vbQuestion Or vbDefaultButton1, "Prospect Status!!")
            Case vbYes
                gconDMIS.Execute ("UPDate CRIS_PROSPECTS  Set STATUS='O' Where PROSPECTID=" & N2Str2Null(labProspectID))
            Case vbNo
                gconDMIS.Execute ("UPDate CRIS_PROSPECTS  Set STATUS='I' Where PROSPECTID=" & N2Str2Null(labProspectID))
        End Select
    Else
        '
        '********RESET THE VARIABLE
        SQL_STATEMENT = ""
        '*************************
    End If
    Dim SQL                                                           As String
    SQL = " INSERT INTO SMIS_SalesOrder("
    SQL = SQL & " ProspectID, SO_No, VI_NO, VDR_NO, Code, DEALER_TYPE, CustName,"
    SQL = SQL & " InvoicedDate, Deyt, HomeTelNo, HomeAddress, OfficeAdd, OfficeTelNo,"
    SQL = SQL & " BirthDate, Spouse, Person, Posisyon, TIN, CTCNo, IssuedAt, IssuedOn,"
    SQL = SQL & " Model, ModelDescription, ProdNo, ConductionSticker, EngineNo, FrameNo,"
    SQL = SQL & " Vino, Plate_No, IGNKEY_NO, Color, Type, Term, FinancingCo, BankTerm,"
    SQL = SQL & " AdditionalInfo, DownPayment, BalToFinanced, MonthsAmort, NetMoAmort,"
    SQL = SQL & " AOR, SalesPrice, NetSalesPrice, Freight, Insurance, LTORegFee, CHMOFee,"
    SQL = SQL & " Accessories, Tax, OthersDesc, Others, Discount, Total, GMI, RPPD, FreeBies,"
    SQL = SQL & " SalesAE, PreparedBy, checkedBy, SalesApproved, SalesDispatcher, DateReleased,"
    SQL = SQL & " Insured, ModeOfPayment, DownpaymentRate, Terms, AutoYesNo, Certific8, Purchaser,"
    SQL = SQL & " ReasonCancel, D_S, STATUS, SOSTATUS, INSUREDDATE, DELIVERY_ADDRESS, DELIVERY_INSTRUCTION, DEBITMEMO,"
    SQL = SQL & " CREDITMEMO , "
    SQL = SQL & " InsuranceCompany, ACCOUNTTYPE)"
    SQL = SQL & " SELECT ProspectID, SO_No, NULL, NULL, "
    SQL = SQL & " Code, DEALER_TYPE, CustName, NULL, "
    SQL = SQL & " Deyt, HomeTelNo, HomeAddress, "
    SQL = SQL & " OfficeAdd, OfficeTelNo, BirthDate, "
    SQL = SQL & " Spouse, Person, Posisyon, TIN, "
    SQL = SQL & " CTCNo, IssuedAt, IssuedOn, "
    SQL = SQL & " Model, ModelDescription, ProdNo, ConductionSticker, EngineNo, "
    SQL = SQL & " FrameNo, Vino, Plate_No, IGNKEY_NO, Color, Type, Term, "
    SQL = SQL & " FinancingCo, BankTerm, AdditionalInfo, DownPayment, "
    SQL = SQL & " BalToFinanced, MonthsAmort, NetMoAmort, AOR, SalesPrice, "
    SQL = SQL & " NetSalesPrice, Freight, Insurance, LTORegFee, CHMOFee, Accessories, Tax, "
    SQL = SQL & " OthersDesc, Others, Discount, Total, GMI, RPPD, FreeBies, "
    SQL = SQL & " SalesAE, PreparedBy, checkedBy, SalesApproved, SalesDispatcher, "
    SQL = SQL & " NULL, Insured, ModeOfPayment, DownpaymentRate, Terms, AutoYesNo, Certific8, Purchaser, "
    SQL = SQL & " NULL,NULL, NULL, SOSTATUS, INSUREDDATE, DELIVERY_ADDRESS, DELIVERY_INSTRUCTION, NULL, "
    SQL = SQL & " NULL, InsuranceCompany, ACCOUNTTYPE FROM SMIS_SalesOrder WHERE ID=" & labid

    gconDMIS.Execute SQL
    '    '*********NEW LOG AUDIT***********
        'SQL_STATEMENT = SQL
        'NEW_LogAudit "A", "SALES ADMIN SALES ORDER", SQL_STATEMENT, labid, "", "INVOICE No:" & txtVINO, "", ""
    '    '********************************

    gconDMIS.Execute ("delete from SMIS_PDI_DET WHERE VI_NO=" & N2Str2Null(txtVINO))
    gconDMIS.Execute ("delete from SMIS_PDI_HDR WHERE VI_NO=" & N2Str2Null(txtVINO))
    gconDMIS.Execute ("DELETE from SMIS_MRRINV_DETAIL WHERE IGNKEYNO=" & N2Str2Null(rsInvoice!IGNKEY_NO))

    txtVehicleDateReleased = ""
    rsRefresh
    rsInvoice.Find "id = " & labid
    StoreMemVars
    ShowHidePictureBox2 picCancelReason, False
End Sub

Private Sub cmdCancelMultiple_Click()
    MULTIPLEVI = False
    ShowHidePictureBox2 picMultipleSO, False
End Sub

Private Sub cmdCancelReason_Click(Index As Integer)
    ShowHidePictureBox2 picCancelReason, False

End Sub

Private Sub cmdCancelRelease_Click(Index As Integer)
    ShowHidePictureBox2 picRelease, False, picAdds
End Sub

Private Sub cmdCancelViewVehicles_Click(Index As Integer)
    ShowHidePictureBox2 picViewVehicles, False
End Sub

Private Sub cmdCloseMultiple_Click()
    cmdCancelMultiple_Click
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "VEHICLES INVOICING") = False Then Exit Sub
    On Error GoTo ErrorCode:
    AddorEdit = "EDIT"
    picHeader.Enabled = True
    cmdSelectVehicles.Enabled = True
    picCustomerInformation.Enabled = True
    picTinInfo.Enabled = True
    picPrintingDetails.Enabled = True
    picTerms.Enabled = True
    picViewAccessories.Enabled = True
    picVehiclesDetail.Enabled = True
    cboSalesOrderNo.Enabled = False
    picAdds.Visible = False
    picSaves.Visible = True
    fraPlateno.Enabled = True
    dtDateInvoiced.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdEditCustInfo_Click()
    Set frmCustomerInfo = New frmAllCustomer
    frmCustomerInfo.AddEditCustomer (txtCusCode.Text)
    'frmCustomerInfo.Show 1
    frmCustomerInfo.Show
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error GoTo ErrorCode:
    frmSMIS_SearchVehicleInvoice.Show
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdFirst_Click()
    On Error GoTo ErrorCode:

    If rsInvoice.BOF Then
        ShowFirstRecordMsg
    Else


        rsInvoice.MoveFirst
        StoreMemVars
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdLast_Click()

    On Error GoTo ErrorCode:

    If rsInvoice.EOF Then
        ShowLastRecordMsg
    Else

        rsInvoice.MoveLast
        StoreMemVars
    End If






    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdNext_Click()

    On Error GoTo ErrorCode:

    rsInvoice.MoveNext
    If rsInvoice.EOF Then
        rsInvoice.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars






    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdOkMaterials_Click()
    If cboAccessories = "" Then: cboAccessories.SetFocus: Exit Sub
    If NumericVal(txtAccRate.Text) = 0 Then: txtAccRate.SetFocus: MsgBox "Accessories Amount must not be Zero", vbOKOnly, "INFORMATION": Exit Sub
    If NumericVal(txtAccQty.Text) = 0 Then: txtAccQty.SetFocus: MsgBox "Accessories Quantity must not Zero", vbOKOnly, "INFORMATION": Exit Sub

    Dim ijx                                                           As Integer
    Dim lst                                                           As ListItem


    On Error Resume Next
    If labEDITDetail = True And Not (lvAccesories.SelectedItem Is Nothing) Then
        lvAccesories.ListItems.Remove (lvAccesories.SelectedItem.Index)

    Else
        ijx = CheckListItem(lvAccesories, cboAccessories)
        If ijx <> -1 Then
            If MsgBox("Free Beeies with Such code Already Exists" & vbCrLf & "Do You Want to Update It", vbYesNo Or vbExclamation Or vbDefaultButton1, App.TITLE) = vbYes Then
                lvAccesories.ListItems.Remove (ijx)
            Else
                ShowHidePictureBox2 picAccessories, False
                Exit Sub
            End If
        End If
    End If

    'UPDATED BY: JUN CEDRON
    'DATE UPDATED:AUGUST 05, 2008
    'DESCRIPTION: FOR ACCESSORIES WITH CHARGE TO CUSTOMER

    If COMPANY_CODE = "HAS" Then
        Dim varISFREE                                                 As Integer
        varISFREE = 1

        If chISFREE.Value = 1 Then
            varISFREE = 0
        End If
    End If


    If COMPANY_CODE = "HAS" Then
        Set lst = lvAccesories.ListItems.Add(, , cboAccessories)
        Call lst.ListSubItems.Add(, , txtAccQty.Text)
        Call lst.ListSubItems.Add(, , txtAccRate.Text)
        Call lst.ListSubItems.Add(, , txtAccAmount.Text)
        Call lst.ListSubItems.Add(, , labAccID)
        Call lst.ListSubItems.Add(, , varISFREE)
    Else
        Set lst = lvAccesories.ListItems.Add(, , cboAccessories)
        Call lst.ListSubItems.Add(, , txtAccQty.Text)
        Call lst.ListSubItems.Add(, , txtAccRate.Text)
        Call lst.ListSubItems.Add(, , txtAccAmount.Text)
        Call lst.ListSubItems.Add(, , labAccID)
    End If
    '
    labEDITDetail = "FALSE"

    Dim TotalAccAmount                                                As Double
    Dim i                                                             As Integer
    Dim SQL                                                           As String
    Dim IGNKEYNO                                                      As String
    IGNKEYNO = txtVehicleConductionSticker

    gconDMIS.Execute "DELETE FROM SMIS_MRRINV_DETAIL WHERE IgnKeyNo =" & N2Str2Null(IGNKEYNO)
    TotalAccAmount = "0.00"

    'UPDATED BY: JUN CEDRON
    'DATE UPDATED:AUGUST 05, 2008
    'DESCRIPTION: FOR ACCESSORIES WITH CHARGE TO CUSTOMER

    If COMPANY_CODE = "HAS" Then
        For i = 1 To lvAccesories.ListItems.Count
            SQL = "INSERT INTO SMIS_MRRINV_DETAIL (IgnKeyNo,Description,QTY, COST,IsFree)values( "
            SQL = SQL & N2Str2Null(IGNKEYNO) & " , "
            SQL = SQL & N2Str2Null(lvAccesories.ListItems(i).Text) & " , "
            SQL = SQL & NumericVal(lvAccesories.ListItems(i).ListSubItems(1).Text) & ", "
            SQL = SQL & NumericVal(lvAccesories.ListItems(i).ListSubItems(2).Text) & ","
            SQL = SQL & NumericVal(lvAccesories.ListItems(i).ListSubItems(5).Text) & ")"
            TotalAccAmount = TotalAccAmount + (NumericVal(lvAccesories.ListItems(i).ListSubItems(1).Text)) * NumericVal(lvAccesories.ListItems(i).ListSubItems(2).Text)
            gconDMIS.Execute SQL
        Next
        SQL_STATEMENT = SQL
    Else
        For i = 1 To lvAccesories.ListItems.Count
            SQL = "INSERT INTO SMIS_MRRINV_DETAIL (IgnKeyNo,Description,QTY, COST,IsFree)values( "
            SQL = SQL & N2Str2Null(IGNKEYNO) & " , "
            SQL = SQL & N2Str2Null(lvAccesories.ListItems(i).Text) & " , "
            SQL = SQL & NumericVal(lvAccesories.ListItems(i).ListSubItems(1).Text) & ", "
            SQL = SQL & NumericVal(lvAccesories.ListItems(i).ListSubItems(2).Text) & ",  1)"
            TotalAccAmount = TotalAccAmount + (NumericVal(lvAccesories.ListItems(i).ListSubItems(1).Text)) * NumericVal(lvAccesories.ListItems(i).ListSubItems(2).Text)
            gconDMIS.Execute SQL
        Next
        SQL_STATEMENT = SQL
    End If

    'UPDATED BY: RDC AUG. 26, 2008
    'THIS IS FOR THE NEW LOG AUDIT
    '**********************************************************************************************
    If COMPANY_CODE = "HAS" Then
        NEW_LogAudit "A", "VEHICLE INVOICING", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtVehicleConductionSticker), "IgnKeyNo", "SMIS_MRRINV_DETAIL"), "", "Invoice No:" & txtVINO, "", ""
    Else
        NEW_LogAudit "A", "VEHICLE INVOICING", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtVehicleConductionSticker), "IgnKeyNo", "SMIS_MRRINV_DETAIL"), "", "Invoice No:" & txtVINO, "", ""
    End If
    '**********************************************************************************************

    txtTotalAccesories = TotalAccAmount
    txtFinAccessories = TotalAccAmount
    txtCashAccessories = TotalAccAmount
    ShowHidePictureBox2 picAccessories, False
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "VEHICLES INVOICING") = False Then Exit Sub

    If IsDate(rsInvoice!DateReleased) = False Then
        MsgBox "Vehicle Not Yet Released, Transaction Cannot Be Posted With Out Releasing", vbExclamation
        Exit Sub
    End If

    If MsgBox("Are you Sure You Want to Post this Transaction", vbInformation + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo ErrorCode:
    If Null2String(rsInvoice!IGNKEY_NO) = "" Then
        MsgBox " Transaction Cannot Be Posted With Out Valid Vehicle Details", vbExclamation
        Exit Sub
    End If
    cmdCancelCO.Enabled = False
    gconDMIS.Execute ("UPDate SMIS_SalesOrder  Set Status='P' Where ID=" & labid)
    rsRefresh
    rsInvoice.Find ("ID=" & labid)
    StoreMemVars

    MessagePop InfoOk, "POSTED", "RECORD SUCESSFULLY POSTED", 1000, 2
    LogAudit "P", "VEHICLE INVOICE", "NAME" & txtCustName & " VEHICLE:" & txtVehicleVinNo
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:
    rsInvoice.MovePrevious
    If rsInvoice.BOF Then
        rsInvoice.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "VEHICLES INVOICING") = False Then Exit Sub
    PRODUCTNO = txtVehicleProdNo.Text
    CUSCODE = txtCusCode.Text
    On Error Resume Next
    '*********RYAN CULAWAY JUNE 13 2008
    'TO DISABLED THE Authorization button and PlateEndingRequest
    'If Not COMPANY_CODE = "HAS" Then '''THIS IS FOR ABAD SANTOS ONLY
    ''    frmSMIS_Report_Print.cmdAuthorization.Enabled = False
    '    frmSMIS_Report_Print.cmdPlateEndingRequest.Enabled = False
    '    frmSMIS_Report_Print.cmdRequestVehicleRegistration.Enabled = False
    'End If
    '*********************************
    
    'UPDATED BY: JUN----------------------------------------------------------
    'DATE UPDATED: 11-20-2008
    If COMPANY_CODE = "HAS" Then
        frmSMIS_Report_Print_HAS.VI_NO = Null2String(rsInvoice!VI_NO)
        frmSMIS_Report_Print_HAS.IGNKEYNO = txtVehicleConductionSticker
        frmSMIS_Report_Print_HAS.GM = txtGeneralManager
        frmSMIS_Report_Print_HAS.Show 1
    ElseIf VALID_COMPANY_CODE_FORHAI = True Then
        frmSMIS_Report_Print_HAI.VI_NO = Null2String(rsInvoice!VI_NO)
        frmSMIS_Report_Print_HAI.IGNKEYNO = txtVehicleConductionSticker
        frmSMIS_Report_Print_HAI.GM = txtGeneralManager
        frmSMIS_Report_Print_HAI.Show 1
    ElseIf COMPANY_CODE = "HSB" Then
        frmSMIS_Report_Print_HSB.VI_NO = Null2String(rsInvoice!VI_NO)
        frmSMIS_Report_Print_HSB.IGNKEYNO = txtVehicleConductionSticker
        frmSMIS_Report_Print_HSB.GM = txtGeneralManager
        frmSMIS_Report_Print_HSB.Show 1
    ElseIf COMPANY_CODE = "HBK" Then
        frmSMIS_Report_Print_HBK.VI_NO = Null2String(rsInvoice!VI_NO)
        frmSMIS_Report_Print_HBK.IGNKEYNO = txtVehicleConductionSticker
        frmSMIS_Report_Print_HBK.GM = txtGeneralManager
        frmSMIS_Report_Print_HBK.Show 1
    ElseIf (COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGO") Then
        frmSMIS_Report_Print_HGC.VI_NO = Null2String(rsInvoice!VI_NO)
        frmSMIS_Report_Print_HGC.IGNKEYNO = txtVehicleConductionSticker
        frmSMIS_Report_Print_HGC.GM = txtGeneralManager
        frmSMIS_Report_Print_HGC.Show 1
    ElseIf COMPANY_CODE = "HMH" Then
        frmSMIS_Report_Print_HMH.VI_NO = Null2String(rsInvoice!VI_NO)
        frmSMIS_Report_Print_HMH.IGNKEYNO = txtVehicleConductionSticker
        frmSMIS_Report_Print_HMH.GM = txtGeneralManager
        frmSMIS_Report_Print_HMH.Show 1
    Else
        frmSMIS_Report_Print.VI_NO = Null2String(rsInvoice!VI_NO)
        frmSMIS_Report_Print.IGNKEYNO = txtVehicleConductionSticker
        frmSMIS_Report_Print.Show 1
    End If
    'UPDATED BY: JUN----------------------------------------------------------
End Sub

Private Sub cmdRefresh_Click()
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset

    If Not (rsInvoice.EOF Or rsInvoice.BOF) Then
        rsRefresh
        rsInvoice.Find ("ID=" & labid)
        StoreMemVars
    End If

    '    If labInvoiceStatus.Caption = "**RELEASED**" Then
    '
    '        SQL = "SELECT Released from SMIS_MRRINV_table where datereleased is not null and released=0 and ignkey='" & txtVehicleConductionSticker.Text & "'"
    '
    '        Set rs = New ADODB.Recordset
    '        Set rs = gconDMIS.Execute(SQL)
    '
    '        If Not rs.EOF And Not rs.BOF Then
    '            gconDMIS.Execute "UPDATE SMIS_MRRINV_table set released=1 where ignkey='" & txtVehicleConductionSticker & "'"
    '        End If
    '    End If
    '    MsgBox "All Information has been refresh.", vbInformation, "Information"
End Sub

Private Sub cmdRelease_Click()
    If Function_Access(LOGID, "Acess_Post", "VEHICLES INVOICING") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox("Are you Sure You Want to Release This Vehicle", vbInformation + vbYesNo) = vbNo Then Exit Sub
    txtRelease_VDR = (GenerateCode("SMIS_SALESORDER", "VDR_NO", "000000"))
    txtRelease_Date = Format(Now, "MM/DD/YYYY")
    txtRelease_Time.Value = TimeValue(LOGTIME)
    ShowHidePictureBox2 picRelease, True, picAdds
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdReleaseVehicle_Click()
    On Error GoTo ErrorCode:
    Dim SQL                                                           As String
    Dim ColorCode                                                     As String
    Dim rsCusVeh1                                                     As ADODB.Recordset


    If Null2String(rsInvoice!IGNKEY_NO) = "" Then
        MsgBox " Transaction Cannot Be Posted With Out Valid Vehicle Details", vbExclamation
        Exit Sub
    End If

    If LTrim(RTrim(txtRelease_VDR)) = "" Then
        ShowIsRequiredMsg " Vehicle Delivery Number"
        txtRelease_VDR.SetFocus
        Exit Sub
    End If
    If IsDate(txtRelease_Date) = False Then
        ShowIsRequiredMsg " Release Date"
        txtRelease_Date.SetFocus
        Exit Sub
    End If

    If DateDiff("d", txtRelease_Date, dtDateInvoiced) > 0 Then
        MsgBox "Release Date Is Less Than Invoice Date! ", vbCritical
        Exit Sub
    End If
    Dim lng                                                           As Integer
    lng = gconDMIS.Execute("select Count(*) from SMIS_SALESORDER WHERE VDR_NO=" & N2Str2Null(LTrim(RTrim(txtRelease_VDR)))).Fields(0).Value

    If txtVehicleMake.Text = "" Then
        ShowIsRequiredMsg ("Make Cannot be Blank")
        Exit Sub
    End If
    
    If lng >= 1 Then
        MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Delivery Reciept Number Already Exist"
        Exit Sub
    End If

    'SMIS_MRRINV_TABLE IS REAL TABLE SMIS_MRRINV IS VIEW FOR CANCELLED

    'gconDMIS.Execute "UPDATE SMIS_MRRINV SET ISTATUS='R' ,RELEASED=1, CUSTOMERCODE= " & N2Str2Null(txtCusCode) & " , PROSPECTID= " & labProspectID & " , DATERELEASED=" & N2Date2Null(txtVehicleDateReleased) & ",  INVOICEDDATE=" & N2Date2Null(dtDateInvoiced) & "," & " VI_NO=" & N2Str2Null(txtVINO) & " WHERE ignkey= '" & LTrim(RTrim(txtVehicleConductionSticker)) & "'"

    If LTrim(RTrim(Null2String(rsInvoice!IGNKEY_NO))) <> txtVehicleConductionSticker Then
        MsgBox txtVehicleConductionSticker
    End If
    SQL_STATEMENT = "UPDATE SMIS_SALESORDER SET STATUS='P', VDR_NO='" & txtRelease_VDR & "', DATERELEASED='" & txtRelease_Date & " " & TimeValue(txtRelease_Time.Value) & "' WHERE ID=" & labid
    gconDMIS.Execute SQL_STATEMENT
    '**********************************
    NEW_LogAudit "RD", "SALES ADMIN SALES ORDER", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtVINO), "VI_NO", "SMIS_SALESORDER"), "", "INVOICE No: " & txtVINO, "", ""
    '**********************************

    gconDMIS.Execute "UPDATE CRIS_PROSPECTS SET INVOICENO='" & txtVINO & "', LOGCLOSINGDATE=" & N2Date2Null(txtRelease_Date) & " , STATUS='C'  WHERE PROSPECTID=" & labProspectID

    gconDMIS.Execute "UPDATE SMIS_MRRINV_TABLE SET ISTATUS='R' ,RELEASED=1, CUSTOMERCODE= " & N2Str2Null(txtCusCode) & " , PROSPECTID= " & labProspectID & " , DATERELEASED=" & N2Date2Null(txtRelease_Date) & ",  INVOICEDDATE=" & N2Date2Null(dtDateInvoiced) & "," & " VI_NO=" & N2Str2Null(txtVINO) & " WHERE UPPER(LTRIM(RTRIM(ignkey)))= '" & UCase(LTrim(RTrim(Null2String(rsInvoice!IGNKEY_NO)))) & "'"


    '    If FormExist("frmSMIS_Trans_VehiclesCheckList") Then
    '        Call frmSMIS_Trans_VehiclesCheckList.SearchByInvoice(txtVINo)
    '    End If


    ColorCode = SetColor(txtVehicleColor)
    Set rsCusVeh1 = New ADODB.Recordset
    Set rsCusVeh1 = gconDMIS.Execute("Select * From CSMS_CUSVEH WHERE UPPER(MAKE) = 'HYUNDAI' AND VCOND_NO='" & txtVehicleConductionSticker & "'")

    If Not (rsCusVeh1.EOF Or rsCusVeh1.BOF) Then
        SQL = " Update CSMS_CUSVEH SET  "
        SQL = SQL & " CUSCDE=" & N2Str2Null(txtCusCode) & ", "
        SQL = SQL & " NIYM= " & N2Str2Null(txtCustName) & ", "
        SQL = SQL & " VIN=" & N2Str2Null(txtVehicleVinNo) & ", "
        SQL = SQL & " PLATE_NO= " & N2Str2Null(txtVehicleConductionSticker) & ", "
        SQL = SQL & " VCOND_NO= " & N2Str2Null(txtVehicleConductionSticker) & ", "
        SQL = SQL & " YER= " & N2Str2Null(txtVehicleYear) & ", "
        SQL = SQL & " MAKE= " & N2Str2Null(txtVehicleMake) & ", "
        SQL = SQL & " MODEL= " & N2Str2Null(txtVehicleModel) & ", "
        SQL = SQL & " MODELCODE= " & N2Str2Null(txtVehicleModelCode) & ", "
        SQL = SQL & " ENGINE= " & N2Str2Null(txtVehicleEngineNo) & ", "
        SQL = SQL & " KMreading= " & N2Str2Null(txtVehicleKMreading) & ", "
        SQL = SQL & " PRODNO= " & N2Str2Null(txtCusCode) & ", "
        SQL = SQL & " SERIAL= " & N2Str2Null(txtVehicleSerialNo) & ", "
        SQL = SQL & " TIN_NUMBER= " & N2Str2Null(txtCusCode) & ", "
        SQL = SQL & " D_SOLD= " & N2Str2Null(dtDateInvoiced) & ", "
        SQL = SQL & " DEL_DATE= " & N2Str2Null(txtRelease_Date) & ", "
        SQL = SQL & " DESCRIPTION= " & N2Str2Null(txtVehicleDescription) & ", "
        SQL = SQL & " CLRCDE= " & N2Str2Null(ColorCode) & ", "
        SQL = SQL & " WAR_CERT=" & N2Str2Null(txtVehicleWarrantyCertifcate) & ", "
        SQL = SQL & " SELLING_DEALER='" & COMPANY_CODE & "'" & ", "
        SQL = SQL & " INVOICENO = " & N2Str2Null(txtVINO)
        SQL = SQL & " WHERE UPPER(MAKE) = 'HYUNDAI' AND VCOND_NO='" & txtVehicleConductionSticker.Text & "'"

        'RESET THE SQL_STATEMENT VARIABLE
        '*************************
        SQL_STATEMENT = ""
        '*************************
        gconDMIS.Execute SQL
        '*************************
        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "CUSTOMER VEHCILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtCusCode), "CUSCDE", "CSMS_CUSVEH", "DETAILS", N2Str2Null(txtVehicleConductionSticker), "VCOND_NO"), "", "SO No: " & cboSalesOrderNo, "", ""
        '*************************
        'LogAudit "A", "CUSTOMER VEHICLE", "CUSTOMER NAME " & txtCustName & " MODEL & txtVehicleModel " & " VIN" & txtVehicleVinNo & " PLATE " & txtVehiclePlateNo
    Else

        SQL = " INSERT INTO CSMS_CUSVEH  ( CUSCDE, NIYM, VIN, PLATE_NO, VCOND_NO,DESCRIPTION, YER, MAKE, "
        SQL = SQL & " MODEL, ENGINE, KMreading, PRODNO, SERIAL, TIN_NUMBER,DEL_DATE, D_SOLD, WAR_CERT,INVOICENO , CLRCDE ,SELLING_DEALER ,MODELCODE) VALUES ( "
        SQL = SQL & N2Str2Null(txtCusCode) & " ,"
        SQL = SQL & N2Str2Null(txtCustName) & " ,"
        SQL = SQL & N2Str2Null(txtVehicleVinNo) & " ,"
        SQL = SQL & N2Str2Null(txtVehicleConductionSticker) & " ,"
        SQL = SQL & N2Str2Null(txtVehicleConductionSticker) & " ,"
        SQL = SQL & N2Str2Null(txtVehicleDescription) & " ,"
        SQL = SQL & N2Str2Null(txtVehicleYear) & " ,"
        SQL = SQL & N2Str2Null(txtVehicleMake) & " ,"
        SQL = SQL & N2Str2Null(txtVehicleModel) & " ,"
        SQL = SQL & N2Str2Null(txtVehicleEngineNo) & " ,"
        SQL = SQL & N2Str2Null(txtVehicleKMreading) & " ,"
        SQL = SQL & N2Str2Null(txtVehicleProdNo) & " ,"
        SQL = SQL & N2Str2Null(txtVehicleSerialNo) & " ,"
        SQL = SQL & N2Str2Null(txtTin) & " ,"
        SQL = SQL & N2Str2Null(txtRelease_Date) & " ,"
        SQL = SQL & N2Str2Null(dtDateInvoiced) & " ,"
        SQL = SQL & N2Str2Null(txtVehicleWarrantyCertifcate) & " ,"
        SQL = SQL & N2Str2Null(txtVINO) & " ,"
        SQL = SQL & N2Str2Null(ColorCode) & ",'" & COMPANY_CODE & "'," & N2Str2Null(txtVehicleModelCode) & " )"

        'RESET THE SQL_STATEMENT VARIABLE
        '*************************
        SQL_STATEMENT = ""
        '*************************
        SQL_STATEMENT = SQL
        NEW_LogAudit "A", "CUSTOMER VEHCILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtCusCode), "CUSCDE", "CSMS_CUSVEH", "DETAILS", N2Str2Null(txtVehicleConductionSticker), "VCOND_NO"), "", "SO No: " & cboSalesOrderNo, "", ""
        '*************************

        'LogAudit "E", "CUSTOMER VEHICLE", "CUSTOMER NAME " & txtCustName & " MODEL & txtVehicleModel " & " VIN" & txtVehicleVinNo & " PLATE " & txtVehiclePlateNo
    End If


    Dim SQLReleased                                                   As String
    Dim RS                                                            As New ADODB.Recordset

    'UPDATE BY : BTT TO VERIFY IF THE TRANSACTION IS ALREADY RELEASE
    SQLReleased = "SELECT RELEASED FROM SMIS_MRRINV_TABLE WHERE RELEASED=0 AND IGNKEY='" & txtVehicleConductionSticker.Text & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQLReleased)

    If Not RS.EOF And Not RS.BOF Then
        gconDMIS.Execute "UPDATE SMIS_MRRINV_table set released=1 where ignkey='" & txtVehicleConductionSticker & "'"
    End If

    MessagePop Star, "Vehicle Released", "Sales Department : Vehicle Sucessfully Released " & vbCrLf & "Service Department : Customer Vehicle Information for Updated.", 4000, 2, 120
    rsInvoice.Requery
    rsInvoice.Find ("ID=" & labid)
    StoreMemVars
    ShowHidePictureBox2 picRelease, False, picAdds
    LogAudit "P", "RELEASED VEHICLE(VEHICLE INVOICING)", "NAME" & txtCustName & " VEHICLE:" & txtVehicleVinNo & " VDR" & txtRelease_VDR & " DATE" & txtRelease_Date
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
'    On Error GoTo Errorcode:

    If txtCusCode = "" Then
        MsgBox " Customer Information Insufficient Please Update Your Sales Order "
        Exit Sub
    End If

    Dim MRRCODE                                                       As String
    Dim rsChkMRRINV                                                   As ADODB.Recordset

    Dim INSURE                                                        As String
    MayoModel = False


    If txtVINO = "" Then
        ShowIsRequiredMsg "Vehicle Invoice"
        Exit Sub
    End If
    If txtVehicleModel = "" Then
        ShowIsRequiredMsg "Invalid Model"
        Exit Sub
    End If

    If MayoModel = True Then
        ShowIsRequiredMsg "Invalid Model"
        Exit Sub
    End If

    If txtVehicleDescription = "" Then
        ShowIsRequiredMsg "Invalid Vehicle Description"
        Exit Sub
    End If

    If txtVehicleConductionSticker = "" Then
        ShowIsRequiredMsg "Invalid Conduction Sticker Number"
        Exit Sub
    End If
    If txtVehicleProdNo = "" Then
        ShowIsRequiredMsg "Invalid ProductNumber"
        Exit Sub
    End If

    If LTrim(RTrim(cboSalesAE)) = "" Then
        ShowIsRequiredMsg "Valid Sales Agent"
        On Error Resume Next
        cboSalesAE.SetFocus
        Exit Sub
    End If

    If cboPaymentTerm.Text = "" Then
        MessagePop InfoWarning, "Invalid Input", "Error in Term!"
        Exit Sub
    End If

    If chkInsurance.Value = 1 Then
        INSURE = "I"
    Else
        INSURE = "N"
    End If


    Dim lng                                                           As Integer
    lng = gconDMIS.Execute("select Count(*) from SMIS_SALESORDER WHERE VI_No=" & N2Str2Null(txtVINO)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Invoice Number Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsInvoice!VI_NO)) <> UCase(txtVINO) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Invoice Number Already Exist"
            Exit Sub
        End If
    End If
    Dim vtxtDate                                                      As String
    Dim vtxtHomeAdd                                                   As String
    Dim vTxtHomeTelNo                                                 As String
    Dim vTxtOfficeAdd                                                 As String
    Dim vTxtOfficeTelNo                                               As String
    Dim vTxtBirthDate                                                 As String
    Dim vTxtSpouse                                                    As String
    Dim vTxtPerson                                                    As String
    Dim vTxtPosisyon                                                  As String
    Dim VtxtTIN                                                       As String
    Dim vTxtCTCNo                                                     As String
    Dim vTxtIssuedAt                                                  As String
    Dim vTxtIssuedOn                                                  As String
    Dim vtxtengineno                                                  As String
    Dim vtxtframeno                                                   As String
    Dim vcbocolor                                                     As String
    Dim vcboType                                                      As String
    Dim vcboFinancingCo                                               As String
    Dim vcboBankTerm                                                  As String
    Dim vcboSalesAE                                                   As String
    Dim vtxtGMI                                                       As Double
    Dim vtxtRPPD                                                      As Double
    Dim vtxtMonths                                                    As Double
    Dim vtxtNetMoAmort                                                As Double
    Dim vtxtVINo                                                      As String
    Dim vtxtVDRNo                                                     As String
    Dim vtxtPlate_No                                                  As String
    Dim vtxtConductionStickerNo                                       As String
    Dim vtxtdatereleased                                              As String
    Dim vtxtPreparedBy                                                As String
    Dim vtxtCheckedBy                                                 As String
    Dim vtxtSalesApproved                                             As String
    Dim vtxtSalesDispatcher                                           As String
    Dim vtxtCashAdditionalInfo                                        As String
    Dim vtxtFinAdditionalInfo                                         As String
    Dim vtxtModeOfPayment                                             As String
    Dim vtxtDownpaymentRate                                           As String
    Dim vtxtTerms                                                     As String
    Dim TIRM                                                          As String
    Dim vtxtCusCode                                                   As String
    Dim vtxtPO                                                        As String
    Dim vtxtAccountType                                               As String
    Dim VTXTSAECODE                                                   As String
    Dim VTXTBANKPO                                                    As String
    Dim VBANKPODATE                                                   As String
    
    vtxtPO = N2Str2Null(SetPo(cboCashModeofPayment))

    Dim SQL                                                           As String
    PURLASTNEYM = N2Str2Null(PURLASTNEYM)
    PURFIRSTNEYM = N2Str2Null(PURFIRSTNEYM)
    PURMIDDLE = N2Str2Null(PURMIDDLE)
    vtxtAccountType = N2Str2Null(SetAccountType(cboAccountType))
    MRRCODE = N2Str2Null(SetMRRCode(txtVehicleProdNo.Text))
    CUSCODE = txtCusCode.Text
    vtxtCusCode = N2Str2Null(txtCusCode.Text)
    vtxtDate = N2Date2Null(dtDateInvoiced.Value)
    vtxtHomeAdd = N2Str2Null(txtHomeAdd)
    vTxtHomeTelNo = N2Str2Null(txtTelephoneHome)
    vTxtOfficeAdd = N2Str2Null(txtOfficeAdd)
    vTxtOfficeTelNo = N2Str2Null(txtTelephoneOffice)
    vTxtBirthDate = N2Date2Null(txtDateBirth.Text)
    vTxtSpouse = N2Str2Null(txtSpouse.Text)
    vTxtPerson = N2Str2Null(txtContactPerson.Text)
    vTxtPosisyon = N2Str2Null(txtPosition.Text)
    VtxtTIN = N2Str2Null(txtTin.Text)
    vTxtCTCNo = N2Str2Null(txtCTCNo.Text)
    vTxtIssuedAt = N2Str2Null(txtIssuedAt.Text)
    vTxtIssuedOn = N2Str2Null(txtIssuedOn.Text)
    vtxtengineno = N2Str2Null(txtVehicleEngineNo.Text)
    vtxtframeno = N2Str2Null(txtVehicleFrameNo.Text)
    vcbocolor = N2Str2Null(txtVehicleColor)
    vcboType = N2Str2Null(cboPurchaseType.Text)
    vcboFinancingCo = N2Str2Null(cboFinFinancingCo.Text)
    vcboBankTerm = N2Str2Null(txtFinBankTerm.Text)
    vcboSalesAE = N2Str2Null(cboSalesAE.Text)
    vtxtGMI = NumericVal(txtFinGMI.Text)
    vtxtRPPD = NumericVal(txtFinRPPD.Text)
    vtxtMonths = N2Str2Zero(txtFinNoOfTermAmort.Text)
    vtxtNetMoAmort = NumericVal(txtFinNetMonthlyAmort.Text)
    vtxtVINo = N2Str2Null(txtVINO.Text)
    vtxtVDRNo = N2Str2Null(labVDRNo)
    vtxtPlate_No = N2Str2Null(txtVehiclePlateNo.Text)
    vtxtConductionStickerNo = N2Str2Null(txtVehicleConductionSticker.Text)
    vtxtdatereleased = N2Date2Null(txtVehicleDateReleased.Text)

    vtxtPreparedBy = N2Str2Null(Mid(txtPreparedBy.Text, 1, 34))
    vtxtCheckedBy = N2Str2Null(Mid(txtCheckedBy.Text, 1, 34))
    vtxtSalesApproved = N2Str2Null(Mid(txtSalesApproved.Text, 1, 34))
    vtxtSalesDispatcher = N2Str2Null(Mid(txtGeneralManager, 1, 34))

    vtxtFinAdditionalInfo = N2Str2Null(infoAdditionalInfo.Text)
    vtxtModeOfPayment = N2Str2Null(cboCashModeofPayment.Text)
    vtxtDownpaymentRate = NumericVal(txtFinDownpaymentRate)
    vtxtTerms = N2Str2Zero(txtFinBankTerm)
    VTXTSAECODE = N2Str2Null(GetSAECode(cboSalesAE))
    VTXTBANKPO = N2Str2Null(txtBankPo.Text)
    VBANKPODATE = N2Str2Null(dtbankcom_po.Value)
    
    If VTXTSAECODE = "" Then
        ShowIsRequiredMsg "Valid Sales Agent"
        On Error Resume Next
        cboSalesAE.SetFocus
        Exit Sub
    End If
    'RDC1205:04292008
    If Not LTrim(RTrim(txtOldCS)) = "" Then
        gconDMIS.Execute ("UPDATE SMIS_MRRINV_TABLE SET PROSPECTID=NULL, customercode=NULL,datereleased=null, invoiceddate=null,IStatus='O', Released=0, WithProsBuyers='N'  WHERE IGNKEY=" & N2Str2Null(LTrim(RTrim(txtOldCS))))
    End If
    gconDMIS.Execute ("UPDATE   SMIS_MRRINV_TABLE SET Istatus='S',  CustomerCode = " & vtxtCusCode & " , ProspectID = " & PROSPECTID & " , Released = 0 ,VI_No=" & vtxtVINo & " WHERE ignkey=" & vtxtConductionStickerNo)

    If UCase(cboPaymentTerm.Text) = "CASH ON DELIVERY" Then
        TIRM = "COD"
    ElseIf UCase(cboPaymentTerm.Text) = "FINANCING" Then
        TIRM = "F"
    ElseIf UCase(cboPaymentTerm.Text) = "BANK PO" Then
        TIRM = "BPO"
    ElseIf UCase(cboPaymentTerm.Text) = "COMPANY PO" Then
        TIRM = "CPO"
    End If

    If TIRM = "COD" Or TIRM = "CPO" Then
        SQL = "UPDATE SMIS_SALESORDER SET " & vbCrLf

        SQL = SQL & " VI_NO = " & vtxtVINo & ", " & vbCrLf
        SQL = SQL & " INVOICEDDATE = " & vtxtDate & "," & vbCrLf
        SQL = SQL & " ACCOUNTTYPE = " & vtxtAccountType & "," & vbCrLf
        SQL = SQL & " CODE = " & vtxtCusCode & ", " & vbCrLf
        SQL = SQL & " CUSTNAME = " & N2Str2Null(txtCustName) & ", " & vbCrLf
        SQL = SQL & " PERSON = " & N2Str2Null(txtContactPerson) & ", " & vbCrLf
        SQL = SQL & " SPOUSE = " & vTxtSpouse & ","
        SQL = SQL & " HOMEADDRESS = " & vtxtHomeAdd & "," & vbCrLf
        SQL = SQL & " HOMETELNO = " & vTxtHomeTelNo & ","
        SQL = SQL & " OFFICEADD = " & vTxtOfficeAdd & "," & vbCrLf
        SQL = SQL & " OFFICETELNO = " & vTxtOfficeTelNo & "," & vbCrLf
        SQL = SQL & " DELIVERY_ADDRESS = " & N2Str2Null(txtDeliveryAddress) & "," & vbCrLf
        SQL = SQL & " BIRTHDATE = " & vTxtBirthDate & "," & vbCrLf
        SQL = SQL & " POSISYON = " & vTxtPosisyon & "," & vbCrLf
        SQL = SQL & " TIN = " & VtxtTIN & "," & vbCrLf
        SQL = SQL & " CTCNO = " & vTxtCTCNo & "," & vbCrLf
        SQL = SQL & " ISSUEDAT = " & vTxtIssuedAt & "," & vbCrLf
        SQL = SQL & " ISSUEDON = " & vTxtIssuedOn & ","
        SQL = SQL & " DELIVERY_INSTRUCTION = " & N2Str2Null(txtDeliveryInstruction) & ","
        SQL = SQL & " ZERORATED = " & chkZeroRate1 & ","
        SQL = SQL & " USERCODE = " & VTXTSAECODE & ","

        SQL = SQL & " MODEL = " & N2Str2Null(txtVehicleModel) & "," & vbCrLf
        SQL = SQL & " PRODNO = " & N2Str2Null(txtVehicleProdNo) & "," & vbCrLf
        SQL = SQL & " ENGINENO = " & vtxtengineno & "," & vbCrLf
        SQL = SQL & " IGNKEY_NO = " & vtxtConductionStickerNo & "," & vbCrLf
        SQL = SQL & " FRAMENO = " & vtxtframeno & "," & vbCrLf
        SQL = SQL & " COLOR = " & vcbocolor & "," & vbCrLf

        SQL = SQL & " SALESPRICE = " & NumericVal(txtCashSalesPrice) & "," & vbCrLf
        SQL = SQL & " DISCOUNT = " & NumericVal(txtCashDiscount) & ", "
        SQL = SQL & " NETSALESPRICE = " & NumericVal(txtCashNetSalesPrice) & "," & vbCrLf
        SQL = SQL & " INSURANCE = " & NumericVal(txtCashInsurance) & "," & vbCrLf

        SQL = SQL & " LTOREGFEE = " & NumericVal(txtCashLTORegFee) & "," & vbCrLf
        SQL = SQL & " ACCESSORIES  = " & NumericVal(txtCashAccessories) & "," & vbCrLf
        SQL = SQL & " TAX = " & NumericVal(txtCashTax) & "," & vbCrLf
        SQL = SQL & " OTHERS = " & NumericVal(txtCashOthers) & "," & vbCrLf
        SQL = SQL & " OTHERSDESC = " & N2Str2Null(txtCashOthersDesc) & ", " & vbCrLf
        SQL = SQL & " TOTAL = " & NumericVal(LAB_TOTAL_CASH) & "," & vbCrLf
        SQL = SQL & " FREIGHT= '" & NumericVal(txtCashFreight) & "', "

        SQL = SQL & " ADDITIONALINFO = " & vtxtFinAdditionalInfo & "," & vbCrLf
        SQL = SQL & " CERTIFIC8 = " & N2Str2Null(txtVehicleWarrantyCertifcate) & "," & vbCrLf
        SQL = SQL & " VDR_NO = " & vtxtVDRNo & "," & vbCrLf
        SQL = SQL & " PLATE_NO = " & vtxtPlate_No & "," & vbCrLf
        SQL = SQL & " SALESAE = " & N2Str2Null(cboSalesAE) & "," & vbCrLf

        SQL = SQL & " PREPAREDBY = " & vtxtPreparedBy & "," & vbCrLf
        SQL = SQL & " CHECKEDBY = " & vtxtCheckedBy & "," & vbCrLf
        SQL = SQL & " SALESAPPROVED = " & vtxtSalesApproved & "," & vbCrLf
        SQL = SQL & " SALESDISPATCHER = " & vtxtSalesDispatcher & ", " & vbCrLf
        SQL = SQL & " INSURED = '" & INSURE & "', " & vbCrLf
        SQL = SQL & " INSUREDDATE = '" & DTPicker2 & "', " & vbCrLf
        SQL = SQL & " MODEOFPAYMENT = " & vtxtPO & ", " & vbCrLf
        SQL = SQL & " BANK_COM_PO = " & VTXTBANKPO & ", " & vbCrLf
        SQL = SQL & " BANK_COM_PO_DATE = " & VBANKPODATE & ", " & vbCrLf
        SQL = SQL & " INSURANCECOMPANY = " & N2Str2Null(cboInsuranceCompany) & ","
        SQL = SQL & " TERM = " & N2Str2Null(TIRM) & vbCrLf
        SQL = SQL & " WHERE ID = " & labid.Caption


    Else
        SQL = "UPDATE SMIS_SALESORDER SET" & vbCrLf
        SQL = SQL & " ZERORATED = " & chkZeroRate2 & ","
        SQL = SQL & " VI_NO = " & vtxtVINo & ", " & vbCrLf
        SQL = SQL & " ACCOUNTTYPE = " & vtxtAccountType & "," & vbCrLf
        SQL = SQL & " CODE = " & vtxtCusCode & ", "
        SQL = SQL & " INVOICEDDATE = " & vtxtDate & ","
        SQL = SQL & " CUSTNAME = " & N2Str2Null(txtCustName) & ", " & vbCrLf
        SQL = SQL & " PERSON = " & N2Str2Null(txtContactPerson) & ", " & vbCrLf
        SQL = SQL & " HOMEADDRESS = " & vtxtHomeAdd & "," & vbCrLf
        SQL = SQL & " HOMETELNO = " & vTxtHomeTelNo & ","
        SQL = SQL & " OFFICEADD = " & vTxtOfficeAdd & "," & vbCrLf
        SQL = SQL & " OFFICETELNO = " & vTxtOfficeTelNo & "," & vbCrLf
        SQL = SQL & " DELIVERY_ADDRESS = " & N2Str2Null(txtDeliveryAddress) & "," & vbCrLf
        SQL = SQL & " BIRTHDATE = " & vTxtBirthDate & ", "
        SQL = SQL & " SPOUSE = " & vTxtSpouse & ","
        SQL = SQL & " POSISYON = " & vTxtPosisyon & ","
        SQL = SQL & " TIN = " & VtxtTIN & ","
        SQL = SQL & " CTCNO = " & vTxtCTCNo & ","
        SQL = SQL & " ISSUEDAT = " & vTxtIssuedAt & ","
        SQL = SQL & " ISSUEDON = " & vTxtIssuedOn & ","
        SQL = SQL & " DELIVERY_INSTRUCTION = " & N2Str2Null(txtDeliveryInstruction) & ","

        SQL = SQL & " MODEL = " & N2Str2Null(txtVehicleModel) & ","
        SQL = SQL & " PRODNO = " & N2Str2Null(txtVehicleProdNo.Text) & ","
        SQL = SQL & " ENGINENO = " & vtxtengineno & ","
        SQL = SQL & " IGNKEY_NO = " & vtxtConductionStickerNo & ", "
        SQL = SQL & " FRAMENO = " & vtxtframeno & ","
        SQL = SQL & " COLOR = " & vcbocolor & ","
        SQL = SQL & " TYPE = " & vcboType & ","

        SQL = SQL & " CERTIFIC8= " & N2Str2Null(txtVehicleWarrantyCertifcate) & ","
        SQL = SQL & " TERM = '" & TIRM & "',"
        SQL = SQL & " FINANCINGCO = " & vcboFinancingCo & ","
        SQL = SQL & " SALESAE = " & vcboSalesAE & ", "
        SQL = SQL & " SALESPRICE = " & NumericVal(txtFinSalesPrice) & ","
        SQL = SQL & " NETSALESPRICE = " & NumericVal(txtFinNetSalesPrice) & ","
        SQL = SQL & " DOWNPAYMENT = " & NumericVal(txtFinDownPayment.Text) & ","
        SQL = SQL & " INSURANCE = " & NumericVal(txtFinInsurance.Text) & ", "
        SQL = SQL & " BALTOFINANCED = " & NumericVal(txtFinBaltoFinanced.Text) & ","
        SQL = SQL & " LTOREGFEE = " & NumericVal(txtFinLTORegFee.Text) & ", "
        SQL = SQL & " CHMOFEE = " & NumericVal(txtFinChattel.Text) & ","
        SQL = SQL & " ACCESSORIES = " & NumericVal(txtFinAccessories.Text) & ", "
        SQL = SQL & " TAX = " & NumericVal(txtFinTax) & "," & vbCrLf
        SQL = SQL & " OTHERS = " & NumericVal(txtFinOthers.Text) & ","
        SQL = SQL & " OTHERSDESC = " & N2Str2Null(txtFinOthersDesc) & ", " & vbCrLf
        SQL = SQL & " GMI = " & vtxtGMI & ","
        SQL = SQL & " RPPD = " & vtxtRPPD & ","
        SQL = SQL & " MONTHSAMORT = " & vtxtMonths & ","
        SQL = SQL & " NETMOAMORT = " & vtxtNetMoAmort & ","
        SQL = SQL & " TOTAL = " & NumericVal(LAB_TOTAL_FIN) & ", "
        SQL = SQL & " DISCOUNT = " & NumericVal(txtFinDiscount) & ", " & vbCrLf
        SQL = SQL & " USERCODE = " & VTXTSAECODE & ","
        SQL = SQL & " ADDITIONALINFO = " & vtxtFinAdditionalInfo & ","
        SQL = SQL & " VDR_NO = " & vtxtVDRNo & ", " & vbCrLf
        SQL = SQL & " PLATE_NO = " & vtxtPlate_No & ", " & vbCrLf
        SQL = SQL & " PREPAREDBY = " & vtxtPreparedBy & "," & vbCrLf
        SQL = SQL & " CHECKEDBY = " & vtxtCheckedBy & "," & vbCrLf
        SQL = SQL & " SALESAPPROVED = " & vtxtSalesApproved & "," & vbCrLf
        SQL = SQL & " SALESDISPATCHER = " & vtxtSalesDispatcher & "," & vbCrLf
        SQL = SQL & " AOR = " & NumericVal(txtFinAOR) & "," & vbCrLf
        SQL = SQL & " BANKTERM = " & vcboBankTerm & "," & vbCrLf
        SQL = SQL & " INSURED = '" & INSURE & "', " & vbCrLf
        SQL = SQL & " INSUREDDATE = '" & DTPicker2 & "', " & vbCrLf
        SQL = SQL & " FREIGHT= '" & NumericVal(txtFinFreight) & "', "
        SQL = SQL & " MODEOFPAYMENT = " & N2Str2Null(txtFinModeofPayment.Text) & ","
        SQL = SQL & " BANK_COM_PO = " & VTXTBANKPO & ", " & vbCrLf
        SQL = SQL & " BANK_COM_PO_DATE = " & VBANKPODATE & ", " & vbCrLf
        SQL = SQL & " DOWNPAYMENTRATE = " & vtxtDownpaymentRate & ","
        SQL = SQL & " INSURANCECOMPANY = " & N2Str2Null(cboInsuranceCompany) & ","
        SQL = SQL & " TERMS = " & vtxtTerms

        SQL = SQL & " WHERE ID = " & labid.Caption

    End If
    gconDMIS.Execute SQL

    ''**UPDATED BY:RDC Aug. 26 2008
    '*************************************************************************************************************************************************************************
    'NEW LOGAUDIT
    '*********************
    SQL_STATEMENT = SQL
    If TIRM = "COD" Or TIRM = "CPO" Then
        NEW_LogAudit "E", "VEHICLE INVOICING", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtVINO), "VI_NO", "SMIS_SALESORDER"), "", "Invoice No: " & txtVINO, "", ""
    Else
        NEW_LogAudit "E", "VEHICLE INVOICING", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtVINO), "VI_NO", "SMIS_SALESORDER"), "", "Invoice No: " & txtVINO, "", ""
    End If
    '*************************************************************************************************************************************************************************
    SQL = ""
    gconDMIS.Execute "update SMIS_LoanIndiv set STATUS='A' where ProspectID=" & labProspectID
    gconDMIS.Execute "UPDATE CRIS_PROSPECTS SET INVOICENO='" & txtVINO & "', LOGCLOSINGDATE=" & N2Date2Null(txtRelease_Date) & " , STATUS='C'  WHERE PROSPECTID=" & labProspectID


    rsRefresh
    rsInvoice.Find ("VI_NO=" & vtxtVINo)
    FillCombo "Select DISTINCT INSURANCECOMPANY FROM SMIS_SALESORDER ORDER BY 1 asc", -1, 0, cboInsuranceCompany
    cmdCancel.Value = True
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdSelectMultiple_Click()
    Dim Item                                                          As ListItem
    Dim SQL                                                           As String
    Dim vtxtDate                                                      As String
    Dim vtxtVINo                                                      As String
    vtxtDate = N2Str2Null(LOGDATE)
    For Each Item In lstMultipleSO.ListItems
        If Item.Checked = True Then
            vtxtVINo = "'" & (GenerateCode("SMIS_SALESORDER", "VI_NO", "000000")) & "'"
            SQL = "UPDATE SMIS_SALESORDER SET " & vbCrLf
            SQL = SQL & " VI_NO = " & vtxtVINo & ", " & vbCrLf
            SQL = SQL & " INVOICEDDATE = " & vtxtDate & " " & vbCrLf
            SQL = SQL & " WHERE ID = " & Item.ListSubItems(4).Text
            gconDMIS.Execute SQL
            SQL = vbNullString
            gconDMIS.Execute "update SMIS_LoanIndiv set STATUS='A' where ProspectID=" & labProspectID

        End If
    Next
    MULTIPLEVI = False
    ShowHidePictureBox2 picMultipleSO, False
    rsRefresh
    StoreMemVars
    MsgBox "All information has been save..", vbInformation, "Information"
End Sub

Private Sub cmdSelectVehicles_Click()
    On Error GoTo ErrorCode
    ' 0       1     2       3       4       5       6       7       8       9      10  11
    'Make, Model , Yeer, Descript, ignkey,prodno,engineno,FrameNo,Vino,SerialNo,color, id
    'If RTrim(LTrim(txtVehicleConductionSticker)) = "" Then
    flex_FillReportView gconDMIS.Execute("SELECT Make, Model , Yeer, Descript, ignkey,prodno,engineno,FrameNo,Vino,SerialNo,color,ModelCode, id , Transmission from SMIS_MRRINV_TABLE WHERE STATUS='P' AND RELEASED=0 "), lvViewVehicles
    lvViewVehicles.Columns(0).Visible = False
    lvViewVehicles.Columns(1).Visible = False
    lvViewVehicles.Columns(5).Visible = False
    lvViewVehicles.Columns(6).Visible = False
    lvViewVehicles.Columns(7).Visible = False
    lvViewVehicles.Columns(9).Visible = False



    ShowHidePictureBox2 picViewVehicles, True
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSelectViewVehicles_Click()
    'Make, Model , Yeer, Descript, ignkey,prodno,engineno,FrameNo,Vino,SerialNo,color,id
    'If Then Exit Sub
    If lvViewVehicles.SelectedRows(0).GroupRow = True Then Exit Sub
    With lvViewVehicles.SelectedRows.Row(0)
        txtVehicleMake = Null2String(.Record(0).Value)
        txtVehicleModel = Null2String(.Record(1).Value)
        txtVehicleYear = Null2String(.Record(2).Value)
        txtVehicleDescription = Null2String(.Record(3).Value)
        txtVehicleConductionSticker = Null2String(.Record(4).Value)
        txtVehicleProdNo = Null2String(.Record(5).Value)
        txtVehicleEngineNo = Null2String(.Record(6).Value)
        txtVehicleFrameNo = Null2String(.Record(7).Value)
        txtVehicleVinNo = Null2String(.Record(8).Value)
        txtVehicleSerialNo = Null2String(.Record(9).Value)
        txtVehicleColor = Null2String(.Record(10).Value)
        txtVehicleModelCode = GetModelCode(Null2String(.Record(1).Value))
        txtMCode = Null2String(.Record(12).Value)
        txtVehicleTransmission = Null2String(.Record(13).Value)
        'ACCESSORIES INFO
        If txtVehicleConductionSticker <> "" Then
            flex_FillListView gconDMIS.Execute("Select Description ,QTY , COST  , QTY * COST, ID from SMIS_MRRINV_DETAIL Where IgnKeyNo =" & N2Str2Null(txtVehicleConductionSticker)), lvAccesories
        End If
    End With

    ShowHidePictureBox2 picViewVehicles, False
End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UnPost", "VEHICLES INVOICING") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If MsgBox("Are you Sure You Want to Un-Post this Transaction", vbInformation + vbYesNo) = vbNo Then Exit Sub

    gconDMIS.Execute ("UPDate SMIS_SalesOrder  Set Status='U' Where ID=" & labid)
    rsRefresh
    rsInvoice.Find ("ID=" & labid)
    StoreMemVars
    MessagePop InfoOk, "UN-POSTED", "RECORD SUCESSFULLY UN-POSTED", 1000, 2
    LogAudit "U", "VEHICLE INVOICE", "NAME" & txtCustName & " VEHICLE:" & txtVehicleVinNo
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdUnReleased_Click()
    If Function_Access(LOGID, "Acess_UnPost", "VEHICLES INVOICING") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If MsgBox(" You Are Unreleasing Vehicles. " & vbCrLf & " Are You Sure ? ", vbQuestion + vbYesNo) = vbYes Then
        SQL_STATEMENT = "UPDATE SMIS_SALESORDER SET DATERELEASED = NULL, Status='U',  VDR_NO=NULL  WHERE  ID=" & labid
        gconDMIS.Execute (SQL_STATEMENT)

        ''**UPDATED BY:RDC Aug. 26 2008
        '**************************************************************************************************************************************************
        NEW_LogAudit "UR", "VEHICLE INVOICING", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtVINO), "VI_NO", "SMIS_SALESORDER"), "", "INVOICE NO:" & txtVINO, "", ""
        '**************************************************************************************************************************************************


        gconDMIS.Execute ("UPDATE SMIS_MRRINV_TABLE SET RELEASED=0, ISTATUS='S' ,DATERELEASED=NULL WHERE IGNKEY=" & N2Str2Null(rsInvoice!IGNKEY_NO))
        gconDMIS.Execute ("UPDATE SMIS_PDI_HDR SET STATUS='O' WHERE VI_NO=" & N2Str2Null(txtVINO))
        'gconDMIS.Execute ("DELETE FROM CSMS_CUSVEH WHERE CUSCDE=" & N2Str2Null(txtCusCode) & " AND VCOND_NO=" & N2Str2Null(txtVehicleConductionSticker))
        gconDMIS.Execute "UPDATE CRIS_PROSPECTS SET INVOICENO=NULL , LOGCLOSINGDATE=NULL , STATUS='C'  WHERE PROSPECTID=" & labProspectID


        rsInvoice.Requery
        rsInvoice.Find ("ID=" & labid)
        StoreMemVars
        MessagePop RecSaveOk, "Un-Released", "Vehicle Sucessfully Un-Released", 1000, 2
        'LogAudit "U", "UN-RELEASED VEHICLE(VEHICLE INVOICING)", "NAME" & txtCustName & " VEHICLE:" & txtVehicleVinNo
        If FormExist("frmSMIS_Trans_VehiclesCheckList") Then
            Call frmSMIS_Trans_VehiclesCheckList.SearchByInvoice(txtVINO)
        End If
    End If
    Exit Sub
    
ErrorCode:
    ShowVBError

End Sub

Private Sub Command1_Click()
    rptFree.WindowTitle = "Freebeeies"
    PrintSQLReport rptFree, SMIS_REPORT_PATH & "Freebies.rpt", "{SMIS_MrrInv_Table.ignkey} = '" & txtVehicleConductionSticker.Text & "' and {SMIS_MrrInv_Detail.IsFree} = true", DMIS_REPORT_Connection, 1
End Sub

Private Sub Command2_Click()
    rptFree.WindowTitle = "Freebeeies Charge"
    PrintSQLReport rptFree, SMIS_REPORT_PATH & "ChargeFreebeies.rpt", "{SMIS_MrrInv_Table.ignkey} = '" & txtVehicleConductionSticker.Text & "' and {SMIS_MrrInv_Detail.IsFree} = False ", DMIS_REPORT_Connection, 1
End Sub

Private Sub Command4_Click()
    '    If AddorEdit = "EDIT" Then
    If Function_Access(LOGID, "ACESS_SYSTEM", "VEHICLES INVOICING") = False Then Exit Sub
    dtDateInvoiced.Enabled = True: dtDateInvoiced.SetFocus
    '   End If
End Sub

Private Sub Command5_Click()
    If MsgBox("Confirm. Do You Want to Delete This Record", vbYesNo + vbInformation) = vbYes Then
        SQL_STATEMENT = "DELETE FROM  SMIS_MrrInv_Detail WHERE ID=" & labAccID
        gconDMIS.Execute SQL_STATEMENT
        
        'UPDATED BY: RDC AUG. 26, 2008
        'THIS IS FOR THE NEW LOG AUDIT
        '**********************************************************************************************
            NEW_LogAudit "XX", "VEHICLE INVOICING", SQL_STATEMENT, labid, "", "Invoice No:" & txtVINO & " - " & cboAccessories, "", labAccID
        '**********************************************************************************************
        
        If txtVehicleConductionSticker <> "" Then
            flex_FillListView gconDMIS.Execute("Select Description ,QTY , COST , QTY * COST, ID from SMIS_MRRINV_DETAIL Where IgnKeyNo =" & N2Str2Null(txtVehicleConductionSticker)), lvAccesories
        Else
            lvAccesories.ListItems.Clear
        End If
        ShowHidePictureBox2 picAccessories, False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF11 Then
        If picAdds.Visible = True Then

            Listview_Loadval lstMultipleSO.ListItems, gconDMIS.Execute("SELECT SO_NO, CUSTNAME, MODELDESCRIPTION ,IGNKEY_NO , ID FROM SMIS_SALESORDER WHERE VI_NO IS NULL AND ISNULL(SOSTATUS,'')<>'C'")
            If lstMultipleSO.ListItems.Count = 0 Then
                MsgBox "There are no Sales Order to Process", vbInformation
            Else
                MULTIPLEVI = True
                ShowHidePictureBox2 picMultipleSO, True
            End If
        End If
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE INVOICING)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "VEHICLE INVOICING")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1


    initMemvars
    rsRefresh


    cboPurchaseType.Clear
    cboPurchaseType.AddItem "1st"
    cboPurchaseType.AddItem "RPL"
    cboPurchaseType.AddItem "ADDL"
    cboPurchaseType.AddItem "TRI"
    cboPurchaseType = "1st"


    With cboPaymentTerm
        .Clear
        .AddItem "Cash On Delivery"
        .AddItem "Financing"
        .AddItem "Bank PO"
        .AddItem "Company PO"
    End With

    cboCashModeofPayment.AddItem "Company PO"
    cboCashModeofPayment.AddItem "Cash"
    cboCashModeofPayment.AddItem "Cheque"

    FillCombo "select Company from SMIS_FinCom order by id asc", -1, 0, cboFinFinancingCo
    FillCombo "select NAME from SMIS_vw_Srep order by id asc", -1, 0, cboSalesAE
    FillCombo "Select COMPANY,ID from SMIS_FinCom  order by COMPANY asc", 1, 0, cboFinFinancingCo
    FillCombo "Select DISTINCT INSURANCECOMPANY FROM SMIS_SALESORDER ORDER BY 1 asc", -1, 0, cboInsuranceCompany

    If COMPANY_CODE = "HAS" Then
        AddColumnHeader "DESCRIPTION,QTY,RATE,AMOUNT,ID,ISFREE", lvAccesories
    Else
        AddColumnHeader "DESCRIPTION,QTY,RATE,AMOUNT", lvAccesories
    End If



    ''select Make, Model , Yeer, Descript, ignkey,prodno,engineno,FrameNo,Vino,SerialNo,color,id  from SMIS_MRRINV
    ReportControlAddColumnHeader lvViewVehicles, "MAKE,MODEL,YEAR,DESCRIPTION, C#, P#, E#,F#,V#,S#, COLOR, #MCODE"
    'lvViewVehicles.GroupsOrder.Add lvViewVehicles.Columns(1)
    'lvViewVehicles.Columns(1).Visible = False
    ReportControlPaintManager lvViewVehicles
    ResizeColumnHeader lvViewVehicles, "8,6,6,20,8,8,8,8,8,8,8,8"
    ResizeColumnHeader lvAccesories, "40,15,15,15"

    picHeader.Enabled = False
    picCustomerInformation.Enabled = False
    picTinInfo.Enabled = False
    picPrintingDetails.Enabled = False
    picTerms.Enabled = False
    picViewAccessories.Enabled = False
    picVehiclesDetail.Enabled = False
    SSTabVDetails.SelectedItem = 0
    
    SEARCH_TAB = "0"

    'UPDATED BY: JUN
    'DATE UPDATED: 08/05/2008
    If COMPANY_CODE <> "HAS" Then
        Command1.Visible = False
        Command2.Visible = False
    End If
StoreMemVars
End Sub

Private Sub frmCustomerInfo_ChangedData(xCUSCODE As String)
    Dim rsCusInfo                                                     As ADODB.Recordset
    Set rsCusInfo = gconDMIS.Execute("Select * from ALL_CUSTOMER WHere CUSCDE=" & N2Str2Null(xCUSCODE))
    If Not (rsCusInfo.EOF Or rsCusInfo.BOF) Then
        If Null2String(rsCusInfo!CUSTYPE) = "P" Then
            txtCusCode = xCUSCODE
            txtCustName = Null2String(rsCusInfo!lastname) + "," + Null2String(rsCusInfo!Firstname) + "." + Null2String(rsCusInfo!MiddleInitial)
            txtDateBirth = Null2String(rsCusInfo!BirthDate)
            txtSpouse = Null2String(rsCusInfo!Spouse)

            If IsNull(rsCusInfo!CITY) = False Then
                txtHomeAdd = Null2String(rsCusInfo!CUSTOMERADD) & ", " & rsCusInfo!CITY
            Else
                txtHomeAdd = Null2String(rsCusInfo!CUSTOMERADD)
            End If
            txtDeliveryAddress = Null2String(rsCusInfo!DELIVERYADDRESS)
            txtOfficeAdd = Null2String(rsCusInfo!CompanyAdd)

            txtTelephoneOffice = Null2String(rsCusInfo!TelephoneNo)
            txtTelephoneHome = Null2String(rsCusInfo!HomePhone)
            txtPosition = Null2String(rsCusInfo!TITLE)

        Else
            txtCusCode = xCUSCODE
            txtCustName = Null2String(rsCusInfo!CUSCOMP)
            txtContactPerson = Null2String(rsCusInfo!lastname) + "," + Null2String(rsCusInfo!Firstname) + "." + Null2String(rsCusInfo!MiddleInitial)
            txtDateBirth = Null2String(rsCusInfo!BirthDate)
            txtSpouse = Null2String(rsCusInfo!Spouse)

            If IsNull(rsCusInfo!CITY) = False Then
                txtHomeAdd = Null2String(rsCusInfo!CUSTOMERADD) & ", " & rsCusInfo!CITY
            Else
                txtHomeAdd = Null2String(rsCusInfo!CUSTOMERADD)
            End If
            txtOfficeAdd = Null2String(rsCusInfo!CompanyAdd)
            txtTelephoneOffice = Null2String(rsCusInfo!TelephoneNo)
            txtTelephoneHome = Null2String(rsCusInfo!HomePhone)
            txtPosition = Null2String(rsCusInfo!TITLE)

        End If

    End If

    Unload frmCustomerInfo
    Set frmCustomerInfo = Nothing
End Sub

Private Sub frmCustomerInfo_ProspectConverted(CustomerCode As String, xGoingWhere As String, PROSPECTID As Long)
    MsgBox "HI"
End Sub

Private Sub labDetails_Click()
    If LTrim(RTrim(LOGCODE)) = "NET" Then
        ShowHidePictureBox2 picNetSpeed, True
    End If
End Sub

Private Sub lstMultipleSO_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
        cmdSelectMultiple.Enabled = True
        Exit Sub
    End If
    For Each Item In lstMultipleSO.ListItems
        If Item.Checked = True Then
            cmdSelectMultiple.Enabled = True
            Exit Sub
        End If
    Next
    cmdSelectMultiple.Enabled = False
End Sub

Private Sub lvAccesories_DblClick()
    If lvAccesories.SelectedItem Is Nothing Then Exit Sub
    If labid = "" Or labid = 0 Then
        Exit Sub
    End If
    Command5.Enabled = True

    'UPDATED BY: JUN
    'DATE UPDATED: 08/05/2008
    If COMPANY_CODE = "HAS" Then
        chISFREE.Visible = True
    Else
        chISFREE.Visible = False
    End If

    ShowHidePictureBox2 picAccessories, True

    labEDITDetail = "TRUE"
    With lvAccesories.SelectedItem
        cboAccessories = .Text
        txtAccQty = NumericVal(.ListSubItems(1).Text)
        txtAccRate = FormatNumber(NumericVal(.ListSubItems(2).Text))
        labAccID = NumericVal(.ListSubItems(4).Text)

        cboAccessories.SetFocus
    End With
End Sub

Private Sub lvViewVehicles_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdSelectViewVehicles_Click
End Sub

Private Sub picNetSpeed_Click()
    If LTrim(RTrim(LOGCODE)) = "NET" Then
        ShowHidePictureBox2 picNetSpeed, False
    End If
End Sub

Private Sub SSTabVDetails_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Index = 2 Then
        cboPaymentTerm_Click
    End If
    'UPDATED BY: JUN
    'DATE UPDATED: 08/05/2008
    If COMPANY_CODE = "HAS" Then
        If Item.Index = 3 Then
            Command1.Visible = True
            Command2.Visible = True
        Else
            Command1.Visible = False
            Command2.Visible = False
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    If labInvoiceStatus.Caption <> "" Then
        If labInvoiceStatus.Visible = True Then
            labInvoiceStatus.Visible = False
        Else
            labInvoiceStatus.Visible = True
        End If
    End If

    If labStatus.Caption <> "" Then
        If labStatus.Visible = True Then
            labStatus.Visible = False
        Else
            labStatus.Visible = True
        End If
    End If
End Sub

Private Sub txtAccAmount_Change()
    UpdateAccessoriesAmount
End Sub

Private Sub txtAccQty_Change()
    UpdateAccessoriesAmount
End Sub

Private Sub txtAccQty_GotFocus()
    If NumericVal(txtAccQty.Text) <= 1 Then txtAccQty = "1"
End Sub

Private Sub txtAccQty_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtAccQty_LostFocus()
    If NumericVal(txtAccQty.Text) <= 1 Then txtAccQty = "1"

End Sub

Private Sub txtAccRate_Change()
    UpdateAccessoriesAmount
End Sub

Private Sub txtCashAccessories_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalCashAmount
End Sub

Private Sub txtCashAccessories_GotFocus()
    If NumericVal(txtCashAccessories.Text) <= 0 Then txtCashAccessories = ""
End Sub

Private Sub txtCashAccessories_LostFocus()
    If NumericVal(txtCashAccessories.Text) <= 0 Then txtCashAccessories = "0.00"
    txtCashAccessories = FormatNumber(txtCashAccessories)
End Sub

Private Sub txtCashDiscount_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalCashAmount
End Sub

Private Sub txtCashDiscount_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCashDiscount_LostFocus()
    If NumericVal(txtCashDiscount.Text) <= 0 Then txtCashDiscount = "0.00"
    txtCashDiscount = FormatNumber(txtCashDiscount)
End Sub

Private Sub txtCashFreight_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalCashAmount
End Sub

Private Sub txtCashFreight_GotFocus()
    If NumericVal(txtCashFreight.Text) <= 0 Then txtCashFreight = ""
End Sub

Private Sub txtCashFreight_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCashFreight_LostFocus()
    If NumericVal(txtCashFreight.Text) <= 0 Then txtCashFreight = "0.00"
    txtCashFreight = FormatNumber(txtCashFreight)
End Sub

Private Sub txtCashInsurance_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalCashAmount
End Sub

Private Sub txtCashInsurance_GotFocus()
    If NumericVal(txtCashInsurance.Text) <= 0 Then txtCashInsurance = ""
End Sub

Private Sub txtCashInsurance_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCashInsurance_LostFocus()
    If NumericVal(txtCashInsurance.Text) <= 0 Then txtCashInsurance = "0.00"
    txtCashInsurance = FormatNumber(txtCashInsurance)
End Sub

Private Sub txtCashLTORegFee_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalCashAmount
End Sub

Private Sub txtCashLTORegFee_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCashLTORegFee_LostFocus()
    If NumericVal(txtCashLTORegFee.Text) <= 0 Then txtCashLTORegFee = "0.00"
    txtCashLTORegFee = FormatNumber(txtCashLTORegFee)
End Sub

Private Sub txtCashNetSalesPrice_GotFocus()
    If NumericVal(txtCashNetSalesPrice.Text) <= 0 Then txtCashNetSalesPrice = ""
End Sub

Private Sub txtCashNetSalesPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCashNetSalesPrice_LostFocus()
    If NumericVal(txtCashNetSalesPrice.Text) <= 0 Then txtCashNetSalesPrice = "0.00"
    txtCashNetSalesPrice = FormatNumber(txtCashNetSalesPrice)
End Sub

Private Sub txtCashOthers_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalCashAmount
End Sub

Private Sub txtCashOthers_GotFocus()
    If NumericVal(txtCashOthers.Text) <= 0 Then txtCashOthers = ""
End Sub

Private Sub txtCashOthers_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCashOthers_LostFocus()
    If NumericVal(txtCashOthers.Text) <= 0 Then txtCashOthers = "0.00"
End Sub

Private Sub txtCashOthersDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtCashSalesPrice_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalCashAmount
End Sub

Private Sub txtCashSalesPrice_GotFocus()
    If NumericVal(txtCashSalesPrice.Text) <= 0 Then txtCashSalesPrice = ""
End Sub

Private Sub txtCashSalesPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCashSalesPrice_LostFocus()
    If NumericVal(txtCashSalesPrice.Text) <= 0 Then txtCashSalesPrice = "0.00"
    txtCashSalesPrice = FormatNumber(txtCashSalesPrice)
End Sub

Private Sub txtCashTax_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalCashAmount
End Sub

Private Sub txtCashTax_GotFocus()
    If NumericVal(txtCashTax.Text) <= 0 Then txtCashTax = ""
End Sub

Private Sub txtCashTax_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCashTax_LostFocus()
    If NumericVal(txtCashTax.Text) <= 0 Then txtCashTax = "0.00"
End Sub

Private Sub txtDeliveryInstruction_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtFilterViewVehicles_Change()
    lvViewVehicles.FilterText = txtFilterViewVehicles.Text
    lvViewVehicles.Populate

    cmdSelectViewVehicles.Enabled = IIf(lvViewVehicles.Rows.Count = 0, False, True)
End Sub

Private Sub txtFinAccessories_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalAmount
End Sub

Private Sub txtFinAccessories_GotFocus()
    If NumericVal(txtFinAccessories.Text) <= 0 Then txtFinAccessories = ""
End Sub

Private Sub txtFinAccessories_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinAccessories_LostFocus()
    If NumericVal(txtFinAccessories.Text) <= 0 Then txtFinAccessories = "0.00"
    txtFinAccessories = FormatNumber(txtFinAccessories)
End Sub

Private Sub txtFinAOR_Change()
    If AddorEdit = "" Then Exit Sub
    cmdAuto_Click

End Sub

Private Sub txtFinAOR_GotFocus()
    If NumericVal(txtCashAccessories.Text) <= 0 Then txtCashAccessories = "0.00"
End Sub

Private Sub txtFinAOR_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinAOR_LostFocus()
    If NumericVal(txtFinAOR.Text) = False Then
        txtFinAOR = "0.000"
    Else
        txtFinAOR = FormatNumber(txtFinAOR, 3)
    End If
End Sub

Private Sub txtFinAOR_Validate(Cancel As Boolean)
    If NumericVal(txtFinAOR) > 100 Then
        Cancel = True
        MessagePop InfoVoid, "INVALID ENTRY", "Please Input Value Less or Equal to 100"
    End If
End Sub

Private Sub txtFinBaltoFinanced_Change()
    If AddorEdit = "" Then Exit Sub
    If NumericVal(txtFinAOR) <= 0 Then Exit Sub
    cmdAuto_Click
End Sub

Private Sub txtFinBankTerm_Change()
    If AddorEdit = "" Then Exit Sub
    txtFinNoOfTermAmort = txtFinBankTerm
    cmdAuto_Click

End Sub

Private Sub txtFinBankTerm_GotFocus()
    If NumericVal(txtFinBankTerm.Text) <= 0 Then txtFinBankTerm = ""
End Sub

Private Sub txtFinBankTerm_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinBankTerm_LostFocus()
    If NumericVal(txtFinBankTerm.Text) <= 0 Then txtFinBankTerm = "0.00"
    txtFinBankTerm = FormatNumber(txtFinBankTerm)
End Sub

Private Sub txtFinChattel_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalAmount
End Sub

Private Sub txtFinChattel_GotFocus()
    If NumericVal(txtFinChattel.Text) <= 0 Then txtFinChattel = ""
End Sub

Private Sub txtFinChattel_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinChattel_LostFocus()
    If NumericVal(txtFinChattel.Text) <= 0 Then txtFinChattel = "0.00"
    txtFinChattel = FormatNumber(txtFinChattel)
End Sub

Private Sub txtFinDiscount_Change()
    If AddorEdit = "" Then Exit Sub
    On Error Resume Next
    ComputebyPert = True
    If COMPANY_CODE = "HNC" Then
        txtFinDownPayment = FormatNumber(NumericVal(txtFinSalesPrice) * (NumericVal(txtFinDownpaymentRate) / 100))
    Else
        txtFinDownPayment = FormatNumber(NumericVal(txtFinNetSalesPrice) * (NumericVal(txtFinDownpaymentRate) / 100))
    End If
    UpdateTotalAmount
End Sub

Private Sub txtFinDiscount_GotFocus()
    If NumericVal(txtFinDiscount.Text) <= 0 Then txtFinDiscount = ""
End Sub

Private Sub txtFinDiscount_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinDiscount_LostFocus()
    If NumericVal(txtFinDiscount.Text) <= 0 Then txtFinDiscount = "0.00"
    txtFinDiscount = FormatNumber(txtFinDiscount)
End Sub

Private Sub txtFinDownPayment_Change()
    If AddorEdit = "" Then Exit Sub
    On Error Resume Next
    UpdateTotalAmount
    If ComputebyPert = False And Tutal > 0 Then
        If COMPANY_CODE = "HNC" Then
            txtFinDownpaymentRate = FormatNumber((NumericVal(txtFinDownPayment) / NumericVal(txtFinSalesPrice)) * 100)
        Else
            If NumericVal(txtFinNetSalesPrice) = 0 Then: Exit Sub
            txtFinDownpaymentRate = FormatNumber((NumericVal(txtFinDownPayment) / NumericVal(txtFinNetSalesPrice)) * 100)
        End If
    End If
End Sub

Private Sub txtFinDownpayment_GotFocus()
    If NumericVal(txtFinDownPayment.Text) <= 0 Then txtFinDownPayment = ""
End Sub

Private Sub txtFinDownpayment_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)

End Sub

Private Sub txtFinDownpayment_LostFocus()
    If NumericVal(txtFinDownPayment.Text) <= 0 Then txtFinDownPayment = "0.00"
    txtFinDownPayment = FormatNumber(txtFinDownPayment)
End Sub

Private Sub txtFinDownpaymentRate_Change()
    If AddorEdit = "" Then Exit Sub
    On Error Resume Next
    UpdateTotalAmount
    If ComputebyPert = True Then
        If COMPANY_CODE = "HNC" Then
            txtFinDownPayment = FormatNumber(NumericVal(txtFinSalesPrice) * (NumericVal(txtFinDownpaymentRate) / 100))
        Else
            txtFinDownPayment = FormatNumber(NumericVal(txtFinNetSalesPrice) * (NumericVal(txtFinDownpaymentRate) / 100))
        End If
    End If

End Sub

Private Sub txtFinDownpaymentRate_GotFocus()

    If NumericVal(txtFinDownpaymentRate.Text) <= 0 Then txtFinDownpaymentRate = ""
    ComputebyPert = True

End Sub

Private Sub txtFinDownpaymentRate_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinDownpaymentRate_LostFocus()
    ComputebyPert = False
    If NumericVal(txtFinDownpaymentRate.Text) <= 0 Then txtFinDownpaymentRate = "0.00"
    txtFinDownpaymentRate = FormatNumber(txtFinDownpaymentRate)
End Sub

Private Sub txtFinFreight_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalAmount
End Sub

Private Sub txtFinFreight_GotFocus()
    If NumericVal(txtFinFreight.Text) <= 0 Then txtFinFreight = ""
End Sub

Private Sub txtFinFreight_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinFreight_LostFocus()
    If NumericVal(txtFinFreight.Text) <= 0 Then txtFinFreight = "0.00"
    txtFinFreight = FormatNumber(txtFinFreight)
End Sub

Private Sub txtFinGMI_GotFocus()
    If NumericVal(txtFinGMI.Text) <= 0 Then txtFinGMI = ""
End Sub

Private Sub txtFinGMI_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinGMI_LostFocus()
    If NumericVal(txtFinGMI.Text) <= 0 Then txtFinGMI = "0.00"
    txtFinGMI = FormatNumber(txtFinGMI)
End Sub

Private Sub txtFinInsurance_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalAmount
End Sub

Private Sub txtFinInsurance_GotFocus()
    If NumericVal(txtFinInsurance.Text) <= 0 Then txtFinInsurance = ""
End Sub

Private Sub txtFinInsurance_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinInsurance_LostFocus()
    If NumericVal(txtFinInsurance.Text) <= 0 Then txtFinInsurance = "0.00"
    txtFinInsurance = FormatNumber(txtFinInsurance)
End Sub

Private Sub txtFinLTORegFee_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalAmount
End Sub

Private Sub txtFinLTORegFee_GotFocus()
    If NumericVal(txtFinLTORegFee.Text) <= 0 Then txtFinLTORegFee = ""
End Sub

Private Sub txtFinLTORegFee_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinLTORegFee_LostFocus()
    If NumericVal(txtFinLTORegFee.Text) <= 0 Then txtFinLTORegFee = "0.00"
    txtFinLTORegFee = FormatNumber(txtFinLTORegFee)
End Sub

Private Sub txtFinNetMonthlyAmort_GotFocus()
    If NumericVal(txtFinNetMonthlyAmort.Text) <= 0 Then txtFinNetMonthlyAmort = ""
End Sub

Private Sub txtFinNetMonthlyAmort_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinNetMonthlyAmort_LostFocus()
    If IsNumeric(txtFinNetMonthlyAmort.Text) = True Then
        txtFinNetMonthlyAmort = FormatNumber(txtFinNetMonthlyAmort)
    Else
        txtFinNetMonthlyAmort = "0.00"
    End If
End Sub

Private Sub txtFinNetSalesPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinNetSalesPrice_LostFocus()
    If NumericVal(txtFinNetSalesPrice.Text) <= 0 Then txtFinNetSalesPrice = "0.00"
    txtFinNetSalesPrice = FormatNumber(txtFinNetSalesPrice)
End Sub

Private Sub txtFinNoOfTermAmort_GotFocus()
    If NumericVal(txtFinNoOfTermAmort.Text) <= 0 Then txtFinNoOfTermAmort = ""
End Sub

Private Sub txtFinNoOfTermAmort_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinNoOfTermAmort_LostFocus()
    If NumericVal(txtFinNoOfTermAmort.Text) <= 0 Then txtFinNoOfTermAmort = "0"
End Sub

Private Sub txtFinOthers_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalAmount
End Sub

Private Sub txtFinOthers_GotFocus()
    If NumericVal(txtFinOthers.Text) <= 0 Then txtFinOthers = ""
End Sub

Private Sub txtFinOthers_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinOthers_LostFocus()
    If NumericVal(txtFinOthers.Text) <= 0 Then txtFinOthers = "0.00"
    txtFinOthers = FormatNumber(txtFinOthers)
End Sub

Private Sub txtFinRPPD_GotFocus()
    If NumericVal(txtFinRPPD.Text) <= 0 Then txtFinRPPD = ""
End Sub

Private Sub txtFinRPPD_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinRPPD_LostFocus()
    If NumericVal(txtFinRPPD.Text) <= 0 Then txtFinRPPD = "0.00"
    txtFinRPPD = FormatNumber(txtFinRPPD)
End Sub

Private Sub txtFinSalesPrice_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateTotalAmount
    On Error Resume Next
    ComputebyPert = True
    
    If COMPANY_CODE = "HNC" Then
        txtFinDownPayment = FormatNumber(NumericVal(txtFinSalesPrice) * (NumericVal(txtFinDownpaymentRate) / 100))
    Else
        txtFinDownPayment = FormatNumber(NumericVal(txtFinNetSalesPrice) * (NumericVal(txtFinDownpaymentRate) / 100))
    End If

End Sub

Private Sub txtFinSalesPrice_GotFocus()
    If NumericVal(txtFinSalesPrice.Text) <= 0 Then txtFinSalesPrice = ""

End Sub

Private Sub txtFinSalesPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinSalesPrice_LostFocus()
    txtFinSalesPrice = FormatNumber(txtFinSalesPrice)
End Sub

Private Sub txtFinTax_GotFocus()
    If NumericVal(txtFinTax.Text) <= 0 Then txtFinTax = ""
End Sub

Private Sub txtFinTax_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFinTax_LostFocus()
    If NumericVal(txtFinTax.Text) <= 0 Then txtFinTax = "0.00"
End Sub

Private Sub txtReasonCancel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtRelease_Date_GotFocus()
    txtRelease_Date = Format(txtRelease_Date, "MM/DD/YYYY")
End Sub

Private Sub txtRelease_Date_LostFocus()
    txtRelease_Date = Format(txtRelease_Date, "MM/DD/YYYY")
End Sub

Private Sub txtRelease_VDR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtRelease_VDR_LostFocus()
    txtRelease_VDR = Format(txtRelease_VDR, "000000")
End Sub

Private Sub txtVehicleConductionSticker_Change()
    If txtVehicleConductionSticker <> "" Then
        lblVehicleStatus = ""
    End If
End Sub

Private Sub txtVehicleEngineNo_LostFocus()
    txtVehicleEngineNo = UCase(txtVehicleEngineNo.Text)
End Sub

Private Sub txtVehicleFrameNo_LostFocus()
    txtVehicleFrameNo = UCase(txtVehicleFrameNo.Text)
End Sub

Private Sub txtVINo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtVINo_LostFocus()
    txtVINO = Format(txtVINO, "000000")
End Sub

