VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCRIS_EntryMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   13530
   Begin VB.PictureBox Picture2 
      Height          =   8010
      Left            =   45
      ScaleHeight     =   7950
      ScaleWidth      =   1845
      TabIndex        =   59
      Top             =   405
      Width           =   1905
      Begin VB.CommandButton Command11 
         Caption         =   "Sales Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         MaskColor       =   &H0000FFFF&
         TabIndex        =   91
         Top             =   2160
         Width           =   1770
      End
      Begin VB.PictureBox picLogBox 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   2160
         Left            =   45
         ScaleHeight     =   2130
         ScaleWidth      =   1770
         TabIndex        =   81
         Top             =   2700
         Width           =   1800
         Begin VB.CommandButton Command2 
            Caption         =   "Preview Logs"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            MaskColor       =   &H0000FFFF&
            TabIndex        =   86
            Top             =   1590
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Log A Call"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            MaskColor       =   &H0000FFFF&
            TabIndex        =   85
            Top             =   390
            Width           =   1590
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Log Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            MaskColor       =   &H0000FFFF&
            TabIndex        =   84
            Top             =   1185
            Width           =   1590
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Log Visit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            MaskColor       =   &H0000FFFF&
            TabIndex        =   83
            Top             =   780
            Width           =   1590
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   315
            Index           =   1
            Left            =   -90
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   0
            Width           =   5370
            _Version        =   655364
            _ExtentX        =   9472
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Prospect Logs"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.01
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   64
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "AOR Computation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   105
         MaskColor       =   &H0000FFFF&
         TabIndex        =   80
         Top             =   7980
         Width           =   1770
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Send Quotation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         MaskColor       =   &H0000FFFF&
         TabIndex        =   79
         Top             =   1148
         Width           =   1770
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Schedule Test Drive"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         MaskColor       =   &H0000FFFF&
         TabIndex        =   78
         Top             =   664
         Width           =   1770
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add Sales Appointment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         MaskColor       =   &H0000FFFF&
         TabIndex        =   77
         Top             =   1650
         Width           =   1770
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New Prospects"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         MaskColor       =   &H0000FFFF&
         TabIndex        =   76
         Top             =   180
         Width           =   1770
      End
   End
   Begin VB.PictureBox picForm 
      Height          =   7785
      Left            =   60
      ScaleHeight     =   7725
      ScaleWidth      =   13425
      TabIndex        =   7
      Top             =   405
      Visible         =   0   'False
      Width           =   13485
      Begin XtremeReportControl.ReportControl ReportControl1 
         Height          =   7215
         Left            =   45
         TabIndex        =   105
         Top             =   315
         Width           =   2310
         _Version        =   655364
         _ExtentX        =   4075
         _ExtentY        =   12726
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   7215
         Left            =   2400
         ScaleHeight     =   7185
         ScaleWidth      =   5595
         TabIndex        =   60
         Top             =   315
         Width           =   5625
         Begin VB.ComboBox cboVehicleCode 
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
            Left            =   135
            TabIndex        =   92
            Tag             =   "@R"
            Top             =   585
            Width           =   1740
         End
         Begin VB.ComboBox cboColors 
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
            Left            =   135
            TabIndex        =   95
            Tag             =   "@R"
            Top             =   1245
            Width           =   2145
         End
         Begin VB.ComboBox cboAttendingSE 
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
            Left            =   135
            TabIndex        =   97
            Tag             =   "@R"
            Top             =   1920
            Width           =   5205
         End
         Begin VB.ComboBox cboVehicles 
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
            Left            =   1935
            TabIndex        =   94
            Tag             =   "@R"
            Top             =   600
            Width           =   3405
         End
         Begin VB.ComboBox cboClassification 
            Height          =   330
            Left            =   135
            TabIndex        =   100
            Top             =   3825
            Width           =   5205
         End
         Begin VB.TextBox txtnotes 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Left            =   135
            MultiLine       =   -1  'True
            TabIndex        =   101
            Top             =   4485
            Width           =   5355
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "&OK"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3450
            MaskColor       =   &H0000FFFF&
            TabIndex        =   102
            Top             =   6660
            Width           =   960
         End
         Begin VB.CommandButton Command6 
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4470
            MaskColor       =   &H0000FFFF&
            TabIndex        =   103
            Top             =   6660
            Width           =   960
         End
         Begin VB.ComboBox cboSubject 
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
            Left            =   135
            TabIndex        =   99
            Top             =   3165
            Width           =   5205
         End
         Begin VB.ComboBox cboLeadSource 
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
            Left            =   135
            TabIndex        =   98
            Tag             =   "@R"
            Top             =   2565
            Width           =   5205
         End
         Begin MSComCtl2.DTPicker dtInitialInquiry 
            Height          =   360
            Left            =   2505
            TabIndex        =   96
            Top             =   1215
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   4194304
            CalendarTitleForeColor=   16777215
            Format          =   20512769
            CurrentDate     =   39084
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   315
            Index           =   7
            Left            =   -90
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   0
            Width           =   6015
            _Version        =   655364
            _ExtentX        =   10610
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "DETAIL INFORMATION OF PROSPECT"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.01
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   64
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Interested Color:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   11
            Left            =   135
            TabIndex        =   68
            Top             =   1005
            Width           =   1440
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Attending Sales Executive / Sales Consultant"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Index           =   8
            Left            =   135
            TabIndex        =   67
            Top             =   1650
            Width           =   4125
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Inquired For"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   66
            Top             =   360
            Width           =   1680
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            Caption         =   "Notes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   65
            Top             =   4245
            Width           =   510
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Classification"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   64
            Top             =   3585
            Width           =   1110
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Subject"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   63
            Top             =   2955
            Width           =   645
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   1
            Left            =   2475
            TabIndex        =   62
            Top             =   975
            Width           =   450
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lead Source"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   61
            Top             =   2325
            Width           =   1035
         End
      End
      Begin VB.PictureBox picInfoForm 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   7215
         Left            =   8070
         ScaleHeight     =   7185
         ScaleWidth      =   5250
         TabIndex        =   9
         Top             =   315
         Width           =   5280
         Begin VB.PictureBox Picture 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   4020
            Index           =   3
            Left            =   45
            ScaleHeight     =   3990
            ScaleWidth      =   5100
            TabIndex        =   10
            Top             =   3105
            Width           =   5130
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Year"
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
               Index           =   20
               Left            =   15
               TabIndex        =   21
               Top             =   2115
               Width           =   1500
            End
            Begin VB.Label lblProspectYear 
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   1530
               TabIndex        =   20
               Top             =   2115
               Width           =   3555
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Make"
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
               Index           =   19
               Left            =   15
               TabIndex        =   19
               Top             =   1830
               Width           =   1500
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Model"
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
               Index           =   18
               Left            =   15
               TabIndex        =   18
               Top             =   1545
               Width           =   1500
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H8000000D&
               Height          =   270
               Index           =   17
               Left            =   15
               TabIndex        =   17
               Top             =   315
               Width           =   1515
            End
            Begin XtremeShortcutBar.ShortcutCaption CapInfo 
               Height          =   315
               Index           =   10
               Left            =   0
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   0
               Width           =   5100
               _Version        =   655364
               _ExtentX        =   8996
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Prospect Inquiry"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.01
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               VisualTheme     =   3
               ForeColor       =   64
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H8000000D&
               Height          =   225
               Index           =   16
               Left            =   15
               TabIndex        =   15
               Top             =   600
               Width           =   5100
            End
            Begin VB.Label lblProspectCode 
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   1545
               TabIndex        =   14
               Top             =   315
               Width           =   3555
            End
            Begin VB.Label lblProspectDescription 
               BackColor       =   &H00FFFFFF&
               Height          =   690
               Left            =   15
               TabIndex        =   13
               Top             =   840
               Width           =   5070
            End
            Begin VB.Label lblProspectModel 
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   1530
               TabIndex        =   12
               Top             =   1545
               Width           =   3555
            End
            Begin VB.Label lblProspectMake 
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   1530
               TabIndex        =   11
               Top             =   1830
               Width           =   3555
            End
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   315
            Index           =   11
            Left            =   -90
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   5370
            _Version        =   655364
            _ExtentX        =   9472
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "DETAIL INFORMATION OF PROSPECT"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.01
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   64
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   9
         Left            =   0
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   0
         Width           =   13455
         _Version        =   655364
         _ExtentX        =   23733
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "::: Add / Edit Prospects:::"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         ForeColor       =   64
      End
   End
   Begin VB.PictureBox picView 
      Height          =   8070
      Left            =   1935
      ScaleHeight     =   8010
      ScaleWidth      =   11550
      TabIndex        =   0
      Top             =   405
      Width           =   11610
      Begin VB.CommandButton Command3 
         Caption         =   "&Disable"
         Height          =   480
         Left            =   8610
         MousePointer    =   99  'Custom
         TabIndex        =   90
         Top             =   7410
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   480
         Left            =   10080
         MousePointer    =   99  'Custom
         TabIndex        =   75
         Top             =   7410
         Width           =   705
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   480
         Left            =   10815
         MousePointer    =   99  'Custom
         TabIndex        =   74
         Top             =   7410
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   480
         Left            =   9345
         MousePointer    =   99  'Custom
         TabIndex        =   73
         Top             =   7410
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   480
         Left            =   7875
         MousePointer    =   99  'Custom
         TabIndex        =   72
         Top             =   7425
         Width           =   705
      End
      Begin VB.PictureBox picInfoView 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   7215
         Left            =   6270
         ScaleHeight     =   7185
         ScaleWidth      =   5250
         TabIndex        =   28
         Top             =   30
         Width           =   5280
         Begin VB.PictureBox picProspectCard 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   2715
            Left            =   60
            ScaleHeight     =   2685
            ScaleWidth      =   5100
            TabIndex        =   44
            Top             =   2880
            Width           =   5130
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Color"
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
               Left            =   15
               TabIndex        =   57
               Top             =   600
               Width           =   1515
            End
            Begin VB.Label lblProspectColor 
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   1545
               TabIndex        =   56
               Top             =   600
               Width           =   3555
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Classification"
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
               Index           =   11
               Left            =   15
               TabIndex        =   55
               Top             =   2400
               Width           =   1500
            End
            Begin VB.Label lblProspectClassification 
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   1530
               TabIndex        =   54
               Top             =   2400
               Width           =   3555
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Date of Inquiry"
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
               Left            =   15
               TabIndex        =   53
               Top             =   2115
               Width           =   1500
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   "LeadSource"
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
               Left            =   15
               TabIndex        =   52
               Top             =   1830
               Width           =   1500
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Vehicles Inquired"
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
               Left            =   15
               TabIndex        =   51
               Top             =   315
               Width           =   1515
            End
            Begin XtremeShortcutBar.ShortcutCaption CapInfo 
               Height          =   315
               Index           =   3
               Left            =   0
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   0
               Width           =   5100
               _Version        =   655364
               _ExtentX        =   8996
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Prospect Inquiry"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               VisualTheme     =   3
               ForeColor       =   64
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Subject"
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
               Index           =   5
               Left            =   15
               TabIndex        =   49
               Top             =   885
               Width           =   5100
            End
            Begin VB.Label lblProspectVehicle 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   270
               Left            =   1545
               TabIndex        =   48
               Top             =   315
               Width           =   3555
            End
            Begin VB.Label lblProspectSubject 
               BackColor       =   &H00FFFFFF&
               Height          =   690
               Left            =   15
               TabIndex        =   47
               Top             =   1125
               Width           =   5070
            End
            Begin VB.Label lblProspectLeadSource 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   270
               Left            =   1530
               TabIndex        =   46
               Top             =   1830
               Width           =   3555
            End
            Begin VB.Label lblProspectInquiry 
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   1530
               TabIndex        =   45
               Top             =   2115
               Width           =   3555
            End
         End
         Begin VB.PictureBox picStatusCard 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   1500
            Left            =   60
            ScaleHeight     =   1470
            ScaleWidth      =   5085
            TabIndex        =   41
            Top             =   5640
            Width           =   5115
            Begin VB.Label lblProspectStatus 
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
               Height          =   1110
               Left            =   30
               TabIndex        =   43
               Top             =   330
               Width           =   5025
            End
            Begin XtremeShortcutBar.ShortcutCaption CapInfo 
               Height          =   315
               Index           =   5
               Left            =   0
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   0
               Width           =   5085
               _Version        =   655364
               _ExtentX        =   8969
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Status Information"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               VisualTheme     =   3
               ForeColor       =   64
            End
         End
         Begin VB.PictureBox picProfileCard 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   2415
            Left            =   60
            ScaleHeight     =   2385
            ScaleWidth      =   5100
            TabIndex        =   29
            Top             =   390
            Width           =   5130
            Begin VB.Label lblEmail 
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   1440
               TabIndex        =   40
               Top             =   2115
               Width           =   3645
            End
            Begin VB.Label lblContactNo 
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   1440
               TabIndex        =   39
               Top             =   1830
               Width           =   3645
            End
            Begin VB.Label lblAddress 
               BackColor       =   &H00FFFFFF&
               Height          =   690
               Left            =   15
               TabIndex        =   38
               Top             =   1125
               Width           =   5070
            End
            Begin VB.Label lblAccountName 
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Left            =   1455
               TabIndex        =   37
               Top             =   600
               Width           =   3645
            End
            Begin VB.Label lblCustomerName 
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
               ForeColor       =   &H00008000&
               Height          =   270
               Left            =   1455
               TabIndex        =   36
               Top             =   315
               Width           =   3645
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Address"
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
               TabIndex        =   35
               Top             =   885
               Width           =   5070
            End
            Begin XtremeShortcutBar.ShortcutCaption CapInfo 
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   0
               Width           =   5100
               _Version        =   655364
               _ExtentX        =   8996
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Profile"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               VisualTheme     =   3
               ForeColor       =   64
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Customer Name"
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
               Index           =   0
               Left            =   15
               TabIndex        =   33
               Top             =   315
               Width           =   1425
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Account Name"
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
               Height          =   255
               Index           =   1
               Left            =   15
               TabIndex        =   32
               Top             =   615
               Width           =   1425
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Contact No:"
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
               TabIndex        =   31
               Top             =   1830
               Width           =   1410
            End
            Begin VB.Label lblCustDetails 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Email"
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
               TabIndex        =   30
               Top             =   2115
               Width           =   1410
            End
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   315
            Index           =   2
            Left            =   -90
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   0
            Width           =   5370
            _Version        =   655364
            _ExtentX        =   9472
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "DETAIL INFORMATION OF PROSPECT"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   64
         End
      End
      Begin VB.PictureBox picProspect 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   7965
         Left            =   30
         ScaleHeight     =   7935
         ScaleWidth      =   6165
         TabIndex        =   23
         Top             =   30
         Width           =   6195
         Begin XtremeReportControl.ReportControl lvGridProspect 
            Height          =   6645
            Left            =   60
            TabIndex        =   24
            Top             =   1200
            Width           =   6060
            _Version        =   655364
            _ExtentX        =   10689
            _ExtentY        =   11721
            _StockProps     =   64
            BorderStyle     =   4
            AllowColumnRemove=   0   'False
            AllowColumnReorder=   0   'False
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            ShowFooter      =   -1  'True
         End
         Begin VB.OptionButton optShowOnly 
            Caption         =   "Closed Prospect"
            Height          =   315
            Index           =   3
            Left            =   3105
            Style           =   1  'Graphical
            TabIndex        =   93
            Tag             =   "Closed"
            Top             =   420
            Width           =   1470
         End
         Begin VB.OptionButton optShowOnly 
            Caption         =   "All Prospects"
            Height          =   315
            Index           =   1
            Left            =   4740
            Style           =   1  'Graphical
            TabIndex        =   89
            Tag             =   "All"
            Top             =   420
            Width           =   1380
         End
         Begin VB.OptionButton optShowOnly 
            Caption         =   "Active Prospect"
            Height          =   315
            Index           =   2
            Left            =   75
            Style           =   1  'Graphical
            TabIndex        =   88
            Tag             =   "Active"
            Top             =   420
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.OptionButton optShowOnly 
            Caption         =   "Inactive Prospects"
            Height          =   315
            Index           =   0
            Left            =   1380
            Style           =   1  'Graphical
            TabIndex        =   87
            Tag             =   "InActive"
            Top             =   420
            Width           =   1695
         End
         Begin VB.TextBox txtFilterProspect 
            Height          =   345
            Left            =   2280
            TabIndex        =   25
            Top             =   810
            Width           =   3510
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search Key"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   13
            Left            =   1320
            TabIndex        =   27
            Top             =   840
            Width           =   945
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   315
            Index           =   4
            Left            =   0
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            Width           =   6180
            _Version        =   655364
            _ExtentX        =   10901
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Select Prospect From List"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   64
         End
      End
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7785
      Left            =   60
      ScaleHeight     =   7755
      ScaleWidth      =   13455
      TabIndex        =   1
      Top             =   405
      Visible         =   0   'False
      Width           =   13485
      Begin XtremeReportControl.ReportControl lvGridCustomer 
         Height          =   6630
         Left            =   60
         TabIndex        =   6
         Top             =   1110
         Width           =   8040
         _Version        =   655364
         _ExtentX        =   14182
         _ExtentY        =   11695
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         AllowColumnResize=   0   'False
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Height          =   345
         Left            =   1935
         TabIndex        =   4
         Top             =   360
         Width           =   1800
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   3780
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox picInfoSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   2910
         Left            =   8115
         ScaleHeight     =   2880
         ScaleWidth      =   5250
         TabIndex        =   70
         Top             =   345
         Width           =   5280
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   315
            Index           =   8
            Left            =   -90
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   0
            Width           =   5370
            _Version        =   655364
            _ExtentX        =   9472
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "DETAIL INFORMATION OF PROSPECT"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   64
         End
      End
      Begin VB.CommandButton cmdaddnewcustomer 
         Caption         =   "Add New Profile"
         Height          =   345
         Left            =   135
         TabIndex        =   5
         Top             =   360
         Width           =   1770
      End
      Begin VB.TextBox txtFilterProfile 
         Height          =   345
         Left            =   60
         TabIndex        =   2
         Top             =   765
         Width           =   4980
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   6
         Left            =   0
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   0
         Width           =   13455
         _Version        =   655364
         _ExtentX        =   23733
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "::: Select Your Customer:::"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         ForeColor       =   64
      End
      Begin VB.Image ImgSearchProspect 
         Height          =   330
         Left            =   5070
         MousePointer    =   99  'Custom
         ToolTipText     =   "Enter Character(s) In Text Box And Press Enter To Search Record In Database"
         Top             =   765
         Width           =   330
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
      Height          =   405
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13545
      _Version        =   655364
      _ExtentX        =   23892
      _ExtentY        =   714
      _StockProps     =   14
      Caption         =   "Prospecting"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.99
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
   End
   Begin VB.Menu mnuContextTRIAD 
      Caption         =   "ContextMenu1"
      Visible         =   0   'False
      Begin VB.Menu mnuAddCustomer 
         Caption         =   "Add New Customer"
      End
      Begin VB.Menu mnuAddCompany 
         Caption         =   "Add New Company"
      End
      Begin VB.Menu mnuAddProspectCompany 
         Caption         =   "Add New Prospective Company"
      End
      Begin VB.Menu mnuAddProspectClient 
         Caption         =   "Add New Prospective Client"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelected 
         Caption         =   "Edit Selected"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh View"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "Select Current"
      End
   End
End
Attribute VB_Name = "frmCRIS_EntryMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProspectID                               As Long
Dim ProfileID                                As Long
Dim ProfileType                              As String
Dim AcctName                                 As String
Dim CUSCDE As String
Dim WithEvents frmxProfile  As frmCRIS_EntryProfile
Attribute frmxProfile.VB_VarHelpID = -1
Dim WithEvents frmxCustomer As frmALLCustomer
Attribute frmxCustomer.VB_VarHelpID = -1



Private Sub cboVehicleCode_Click()
    cboVehicles.ListIndex = SelectCombo(cboVehicleCode, cboVehicleCode.ItemData(cboVehicleCode.ListIndex), True)
End Sub

Private Sub cboVehicles_Click()
    If cboVehicles.ListIndex = -1 Then: Exit Sub
    Dim temprs                               As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT * From ALL_MODEL WHERE ID=" & cboVehicles.ItemData(cboVehicles.ListIndex))
    If Not (temprs.EOF Or temprs.BOF) Then
        cboVehicleCode.Text = Null2String(temprs!code)
        lblProspectCode = Null2String(temprs!code)
        lblProspectDescription = Null2String(temprs!descript)
        lblProspectModel = Null2String(temprs!Model)
        lblProspectMake = Null2String(temprs!Make)
        lblProspectYear = Null2String(temprs!yeer)


    End If
End Sub

Private Sub cmdAdd_Click()
'On Error GoTo adder:
    'If Not lvGridCustomer.Records.Count = 0 Then
        'With lvGridProspect.SelectedRows.Row(0)
            ProspectID = 0
            ProfileType = vbNullString                        '.Record(11).Value
            ProfileID = 0                                     ' .0Record(12).Value
            AcctName = vbNullString                           '.Record(2).Value
        'End With
    'End If
    
    picView.Visible = False: picForm.Visible = False: picSearch.Visible = True: Picture2.Visible = False
    Set picProfileCard.Container = picInfoSearch
    lblCustomerName = vbNullString
    AcctName = vbNullString
    lblAccountName = vbNullString
    lblAddress = vbNullString
    lblContactNo = vbNullString
    lblEmail = vbNullString

    flex_FillReportView gconDMIS.Execute("Select TOP 100 * from CRIS_vW_AllProfile order by ProfileName, Acctname"), lvGridCustomer, False
    
'Exit Sub
'adder:
'Err.Clear

End Sub





Private Sub cmdaddnewcustomer_Click()
     frmCRIS_EntryProfile.AddProfile
     frmCRIS_EntryProfile.Show
End Sub

Private Sub cmdDelete_Click()
    If MsgBox(" Confirm :: " & vbCrLf & "  Do you Want to Delete this Prospect", vbOKCancel + vbInformation) = vbOK Then
        MessagePop DELETE, " Record Deleted", " Prospect Information Deleted", 1000, 1
    End If
End Sub

Private Sub cmdEdit_Click()
    If ProspectID <= 0 Then
        MessagePop InfoVoid, "Selection Required", "Please Select Your Prospect From The WIndow", 1000, 1
        Exit Sub
    End If
    
Dim temprs As ADODB.Recordset
Set temprs = gconDMIS.Execute("select COUNT(*) FROM CRIS_PROSPECTS WHERE PROSPECTID=" & ProspectID & " and LogClosingDate is Not Null")
    If temprs.Collect(0) > 0 Then
        MsgBox "This Prospect Has Already Been Closed.."
        Exit Sub
    End If


    loadProspect
    picView.Visible = False: picForm.Visible = True: picSearch.Visible = False
    Set picProfileCard.Container = picInfoForm
End Sub
Sub loadProspect()
    Dim temprs                               As ADODB.Recordset
    Set temprs = gconDMIS.Execute("Select * from CRIS_PROSPECTS Where ProspectID=" & ProspectID)
    If Not (temprs.EOF Or temprs.BOF) Then
        cboVehicleCode.Text = Null2String(temprs!VehicleCode)
        cboVehicles.Text = Null2String(temprs!VehicleModel)
        cboColors.Text = Null2String(temprs!Color)
        dtInitialInquiry.Value = Null2Date(temprs!LogInitialInquiry)
        cboAttendingSE.Text = Null2String(temprs!SAE)
        cboLeadSource.Text = Null2String(temprs!LeadSource)
        cboSubject.Text = Null2String(temprs!Subject)
        cboClassification.Text = Null2String(temprs!Classification)
        txtnotes.Text = Null2String(temprs!notes)
    End If


End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub cmdSelect_Click()
    If ProfileID = 0 Or AcctName = vbNullString Or ProfileType = vbNullString Then
        MsgBox " Please Select Selection From the List"
    Else
        picView.Visible = False: picForm.Visible = True: picSearch.Visible = False
        Set picProfileCard.Container = picInfoForm
        LabelIt
    End If
End Sub

'Private Sub cmdSave_Click()
'Dim SQL As String
'
'If AppointmentID <= 0 Then
'    SQL = " INSERT INTO CRIS_CalendarEvents(AppointmentType, ProspectID, PossibleNextVisit, StartDateTime, EndDateTime, Data1) " _
     '          & " VALUES(@AppointmentTYpe, @ProspectID, @PossibleNextVisit, @StartDateTime, @EndDateTime, @Data1) " & vbCrLf & "SELECT @@IDENTITY"
'
'Else
'    SQL = "Update CRIS_CalendarEvents SET " _
     '          & " AppointmentType=@AppointmentTYpe, ProspectID=@ProspectID, PossibleNextVisit=@PossibleNextVisit, StartDateTime=@StartDateTime, EndDateTime=@EndDateTime, Data1=@Data1 Where AppointmentID=@AppointmentID"
'End If
'
'    SQL = Replace(SQL, "@AppointmentID", AppointmentID)
'    SQL = Replace(SQL, "@AppointmentTYpe", "2")
'    SQL = Replace(SQL, "@ProspectID", ProspectID)
'    SQL = Replace(SQL, "@PossibleNextVisit", N2Str2Null(dtNextVisit.Value))
'    SQL = Replace(SQL, "@Data1", N2Str2Null(txtBody.Text))
'    SQL = Replace(SQL, "@StartDateTime", N2Str2Null(DateFromString(CStr(dtStartDate.Value), CStr(dtStartTime))))
'    SQL = Replace(SQL, "@EndDateTime", N2Str2Null(DateFromString(CStr(dtEndDate.Value), CStr(dtEndTime.Value))))
'
'
'    Dim temprs As ADODB.Recordset
'
'     Set temprs = gconDMIS.Execute(SQL)
'
'    If AppointmentID <= 0 Then
'        MessagePop RecSave, "Record Added", "New Sales Appointment Added", 500, 1
'    Else
'        MessagePop RecSaveOk, "Record Saved", "Sales Appointment Updated", 500, 1
'    End If
'    Set temprs = temprs.NextRecordset
'
'    If Not temprs Is Nothing Then
'        AppointmentID = temprs.Collect(0)
'
'    End If
'        gconDMIS.Execute "update CRIS_PROSPECTS SET LogAppointment=" & N2Str2Null(DateFromString(CStr(dtStartDate.Value), CStr(dtStartTime))) & " Where ProspectID=" & ProspectID
'
'    Set temprs = Nothing
'    ShowHide picAppointment.hwnd, False
'    FillGrid
'End Sub



Private Sub Command1_Click()
    cmdAdd_Click
    txtFilterProfile.SetFocus
End Sub

Private Sub Command10_Click()
    If ProspectID <= 0 Then
        MessagePop InfoVoid, "Selection Required", "Please Select Your Prospect From The WIndow", 1000, 1
        Exit Sub
    End If
    Call frmCRIS_LogCall.AddCall(ProspectID, ProfileType, AcctName, ProspectID)
    frmCRIS_LogCall.Show
End Sub

Private Sub Command11_Click()
    If ProspectID <= 0 Then
        MessagePop InfoVoid, "Selection Required", "Please Select Your Prospect From The WIndow", 1000, 1
        Exit Sub
    End If
    frmcris_SalesOrder.AddNewSO ProspectID, ProfileType, AcctName, ProfileID
     frmcris_SalesOrder.Show
     
    If optShowOnly(0).Value = True Then
        FillGrid 1
    ElseIf optShowOnly(1).Value = True Then
        FillGrid 2
    Else
        FillGrid 3
    End If
    
End Sub

Private Sub Command12_Click()
    If ProspectID <= 0 Then
        MessagePop InfoVoid, "Selection Required", "Please Select Your Prospect From The WIndow", 1000, 1
        Exit Sub
    End If
    Call frmCRIS_LogEmail.AddEmail(ProfileID, ProfileType, AcctName, ProspectID)
    frmCRIS_LogEmail.Show
End Sub

Private Sub Command13_Click()
    If ProspectID <= 0 Then
        MessagePop InfoVoid, "Selection Required", "Please Select Your Prospect From The WIndow", 1000, 1
        Exit Sub
    End If
    Call frmCRIS_LogVisit.AddVisit(ProfileID, ProfileType, AcctName, ProspectID)
    frmCRIS_LogVisit.Show
End Sub

Private Sub Command14_Click()
    frmCRIS_Inquiry.Show
End Sub

Private Sub Command16_Click()
    frmCRIS_AOR.Show
End Sub

Private Sub Command2_Click()
    If ProspectID <= 0 Then
        MessagePop InfoVoid, "Selection Required", "Please Select Your Prospect From The WIndow", 1000, 1
        Exit Sub
    End If
    frmCRIS_ViewLog.ShowReport ProspectID, AcctName

    frmCRIS_ViewLog.Show
End Sub

Private Sub Command3_Click()
'If lvGridCustomer.Records.Count = 0 Then
'    MessagePop NoEntry, "Selection Required", "Please Select Prospect From The List", 1000, 1
'Exit Sub
'End If

    If IsNull(lvGridProspect.SelectedRows.Row(0).Record(13)) = False Then
        If MsgBox("The Prospect is Active...Do You Want To Disable This Prospect?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Confirm") = vbYes Then
            GoTo DELETE:
        End If
    Else
        If MsgBox("Do You Want To Disable This Prospect?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Confirm") = vbYes Then
            GoTo DELETE:
        End If
    End If

    Exit Sub
DELETE:
lvGridProspect.SelectedRows.DeleteAll
    MsgBox "DELETE"
    
End Sub

Private Sub Command4_Click()
    picView.Visible = True: picForm.Visible = False: picSearch.Visible = False: Picture2.Visible = True
    Set picProfileCard.Container = picInfoView
End Sub

Private Sub Command5_Click()
    MsgBox ProspectID
End Sub

Private Sub Command6_Click()
'    picView.Visible = True: picForm.Visible = False: picSearch.Visible = False: Picture2.Visible = True
'    Set picProfileCard.Container = picInfoView
cmdAdd_Click
End Sub

Private Sub Command7_Click()
    If ProspectID <= 0 Then
        MessagePop InfoVoid, "Selection Required", "Please Select Your Prospect From The Window", 1000, 1
        Exit Sub
    End If
    Call frmCRIS_SalesAppointment.AddSalesAppointment(ProspectID, AcctName, ProfileType, ProfileID)
    frmCRIS_SalesAppointment.Show
End Sub

Private Sub Command8_Click()
    If ProspectID <= 0 Then
        MessagePop InfoVoid, "Selection Required", "Please Select Your Prospect From The WIndow", 1000, 1
        Exit Sub
    End If
    Call frmCRIS_EntryTestDriveAppointment.AddTestDriveAppointment(ProspectID, AcctName, ProfileType, ProfileID)
    frmCRIS_EntryTestDriveAppointment.Show
End Sub

Private Sub Command9_Click()
    If ProspectID <= 0 Then
        MessagePop InfoVoid, "Selection Required", "Please Select Your Prospect From The Prospect List", 1000, 1
        Exit Sub
    End If

    Call frmCRIS_EntryQuotation.NewQuotation(ProfileID, ProfileType, AcctName, ProspectID)
    frmCRIS_EntryQuotation.Show
End Sub

Private Sub Form_Load()
    InitVars
    CenterMe frmMain, Me, 1
End Sub
Private Sub InitVars()
    Dim temprs                               As ADODB.Recordset
    With lvGridProspect
        .Columns.Add 0, "ID", 0, True
        .Columns.Add 1, "AcctName", 50, True
        .Columns.Add 2, "Vehicle", 100, True
        .Columns(0).Visible = False
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots  ' xtpGridNoLines
        .PaintManager.GroupRowTextBold = True                 ' = vbWhite
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.HideSelection = True
        .PaintManager.HotTracking = True
        .PaintManager.CaptionFont.Bold = True
        '.PaintManager.ColumnStyle = xtpColumnOffice2003
        .Columns(1).FooterAlignment = xtpAlignmentRight
        .Columns(1).FooterText = " Filter"
    End With
    With lvGridCustomer
        .Columns.Add 0, "ID", 0, True
        .Columns.Add 1, "Account Name", 50, True
        .Columns.Add 2, "Profile Name", 100, True
        .Columns(0).Visible = False
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots  ' xtpGridNoLines
        .PaintManager.GroupRowTextBold = True                 ' = vbWhite
        '.PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        '.SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
    End With
    FillGrid 1



    Call FillCombo("Select DataID as ID, " _
                 & " MasterData as [Description] " _
                 & " from CRIS_vw_master_PullDown Where MasterType='Inquiry Type'", 0, 1, cboSubject)

    Call FillCombo("Select DISTINCT 1, COLOR_DESC FROM ALL_COLOR ORDER BY COLOR_DESC", 0, 1, cboColors)
    
    'Call FillCombo("SELECT ID, DESCRIPT from ALL_MODEL", 0, 1, cboVehicles)
    Set temprs = gconDMIS.Execute("SELECT ID as I  , DESCRIPT as D  ,CODE as C from ALL_MODEL")
    
    While Not (temprs.EOF Or temprs.BOF)
    cboVehicleCode.AddItem (temprs!C)
    cboVehicleCode.ItemData(cboVehicleCode.NewIndex) = temprs!i
    
    cboVehicles.AddItem (temprs!D)
    cboVehicles.ItemData(cboVehicles.NewIndex) = temprs!i
    
        temprs.MoveNext
    Wend
    
    Call FillCombo("SELECT ID, LastName + ', ' + FirstName + '.'+ MiddleName from HRMS_EMPINFO WHERE IS_SAE=1 ORDER BY LASTNAME", 0, 1, cboAttendingSE)



    Set temprs = gconDMIS.Execute("Select DataID, MasterData ,MasterType from CRIS_vw_master_PullDown where MasterType IN ('Customer Classification', 'Lead Source', 'Subject')")



    While Not temprs.EOF
        If temprs.Fields("MasterType").Value = "Lead Source" Then
            cboLeadSource.AddItem temprs.Collect(1)
            cboLeadSource.ItemData(cboLeadSource.NewIndex) = temprs.Fields(0).Value
        ElseIf temprs.Fields("MasterType").Value = "Customer Classification" Then
            cboClassification.AddItem temprs.Collect(1)
            cboClassification.ItemData(cboClassification.NewIndex) = temprs.Fields(0).Value
        ElseIf temprs.Fields("MasterType").Value = "Customer Classification" Then
            cboClassification.AddItem temprs.Collect(1)
            cboClassification.ItemData(cboClassification.NewIndex) = temprs.Fields(0).Value
        End If
        temprs.MoveNext

    Wend



End Sub
Sub FillGrid(ProspectState As Integer)
    Dim temprs                               As ADODB.Recordset
    If ProspectState = 1 Then
        Set temprs = gconDMIS.Execute("SELECT  ProspectID, " & _
                                     "AcctName , VehicleModel ," & _
                                     "VehicleModel, LogQuote, LogEmail, LogAppointment, LogTestDrive,  " & _
                                     "LogCall, LogJournal, LogLetter, ProfileType , ProfileID, U_S FROM CRIS_Prospects Where D_S is NULL and LOGCLOSINGDATE IS NULL ")
    ElseIf ProspectState = 2 Then
        Set temprs = gconDMIS.Execute("SELECT  ProspectID, " & _
                                     "AcctName , VehicleModel ," & _
                                     "VehicleModel, LogQuote, LogEmail, LogAppointment, LogTestDrive,  " & _
                                     "LogCall, LogJournal, LogLetter, ProfileType , ProfileID ,U_S FROM CRIS_Prospects Where D_S is Not NULL  ")
    ElseIf ProspectState = 3 Then
        Set temprs = gconDMIS.Execute("SELECT  ProspectID, " & _
                                     "AcctName , VehicleModel ," & _
                                     "VehicleModel, LogQuote, LogEmail, LogAppointment, LogTestDrive,  " & _
                                     "LogCall, LogJournal, LogLetter, ProfileType , ProfileID, U_S FROM CRIS_Prospects ")
Else
    Set temprs = gconDMIS.Execute("SELECT  ProspectID, " & _
                                     "AcctName , VehicleModel ," & _
                                     "VehicleModel, LogQuote, LogEmail, LogAppointment, LogTestDrive,  " & _
                                     "LogCall, LogJournal, LogLetter, ProfileType , ProfileID, U_S FROM CRIS_Prospects WHERE LOGCLOSINGDATE IS NOT NULL ")
End If

    flex_FillReportView temprs, lvGridProspect, False



    'convert(varchar, LogInitialInquiry ,101)as [Date] ,
End Sub




Public Sub NewEvent(pid As Long, strProfilType As String)



    ProfileID = pid
    ProfileType = strProfilType



End Sub

Public Sub SetStartEnd(BeginSelection As Date, EndSelection As Date, AllDay As Boolean)
    Dim StartDate                            As Date
    Dim StartTime                            As Date
    Dim EndDate                              As Date
    Dim EndTime                              As Date

    StartDate = DateValue(BeginSelection)
    StartTime = TimeValue(BeginSelection)
    EndDate = DateValue(EndSelection)
    EndTime = TimeValue(EndSelection)
    If AllDay Then
        dtStartTime.Visible = False
        dtEndTime.Visible = False

        If DateDiff("s", EndTime, 0) = 0 Then
            EndDate = EndDate - 1
        End If
    End If

    dtStartTime.Value = StartTime
    dtEndTime.Value = EndTime
End Sub


Public Sub GetEvent()
    '    Dim temprs                               As ADODB.Recordset
    '    Dim bDatesVisible                        As Boolean
    '    Dim pLabel                               As CalendarEventLabel
    '
    '    Set temprs = gconDMIS.Execute("Select * from CRIS_CalendarEvents where AppointmentID=" & EventCalendar.ID)
    '    cboSubject.Text = Null2String(temprs!Subject)
    '    ProfileId = N2Str2IntZero(temprs!ProspectID)
    '    'ProfileType = temprs!ProfileType
    '    txtBody.Text = EventCalendar.Body
    '    txtLocation.Text = EventCalendar.Location
    '    chkAllDayEvent.Value = IIf(EventCalendar.AllDayEvent, 1, 0)
    '    cboLabel.ListIndex = SelectComboItemData(EventCalendar.Label, cboLabel)
    '    cboVehicles.ListIndex = SelectComboItemData(IIf(IsNull(temprs!VehicleID), 0, temprs!VehicleID), cboVehicles)
    '    cboAttendingSE.ListIndex = SelectComboItemData(N2Str2IntZero(temprs!SAE), cboAttendingSE)
    '    cboColors.ListIndex = SelectComboItemData(N2Str2IntZero(temprs!color), cboColors)
    '    If IsNull(temprs!PossibleNextVisit) = False Then
    '        dtNextVisit.Value = temprs!PossibleNextVisit
    '        chkPossibleNextVisit.Value = 1
    '    End If
    '
    '    Set pLabel = frmCRIS_DashBoard.cCalSales.DataProvider.LabelList.Find(EventCalendar.Label)
    '
    '    If Not pLabel Is Nothing Then
    '        picCtrlColor.BackColor = pLabel.color
    '    End If
    '    Dim i                                    As Long
    '    For i = 0 To cboLabel.ListCount - 1
    '        If cboLabel.ItemData(i) = EventCalendar.Label Then
    '            cboLabel.ListIndex = i
    '            Exit For
    '        End If
    '    Next
    '    SetStartEnd EventCalendar.StartTime, EventCalendar.EndTime, EventCalendar.AllDayEvent
    '    bDatesVisible = EventCalendar.RecurrenceState <> xtpCalendarRecurrenceMaster
    '    cboStartDate.Visible = bDatesVisible
    '    dtStartTime.Visible = bDatesVisible
    '    cboEndDate.Visible = bDatesVisible
    '    dtEndTime.Visible = bDatesVisible
    '    chkAllDayEvent.Visible = bDatesVisible
    '    If bDatesVisible Then
    '        chkAllDayEvent_Click
    '    End If
    '    LabelIt
End Sub



Function IsDateValid(DatePart As String) As Boolean
    IsDateValid = False
    On Error GoTo Error
    Dim dtDate                               As Date
    dtDate = DatePart
    IsDateValid = True
Error:
    Exit Function
End Function



Function DateFromString(DatePart As String, TimePart As String) As Date
    Dim dtDatePart As Date, dtTimePart       As Date
    dtDatePart = DatePart
    dtTimePart = TimePart
    DateFromString = dtDatePart + dtTimePart
End Function
Private Sub cmdOk_Click()
    Dim SQL                                  As String

    If ProspectID <= 0 Then
        SQL = "INSERT INTO CRIS_Prospects ( " & _
              "CUSCDE, VehicleID,VehicleCode, ProfileID, AcctName, ProfileType, LeadSource, VehicleModel,  " & _
              "Color, SAE, Classification, Notes, Subject, LogInitialInquiry) " & _
              "VALUES(@CUSCDE , @VID,@VehicleCode , @ProfileID, @AcctName, @ProfileType, @LeadSource, @VehicleModel,  " & _
              "@Color, @Sae, @Classification, @Notes, @Subject, @LogInitialInquiry) "
    Else

        SQL = "UPDATE CRIS_Prospects SET   " & _
              "CUSCDE =@CUSCDE , VehicleID=@VID, VehicleCode= @VehicleCode, ProfileID=@ProfileID, AcctName=@AcctName, ProfileType=@ProfileType,  " & _
              "LeadSource=@LeadSource, VehicleModel=@VehicleModel, Color=@Color, SAE=@Sae,  " & _
              "Classification=@Classification, Notes=@Notes, Subject=@Subject, LogInitialInquiry=@LogInitialInquiry " & _
              "WHERE ProspectID=@ProspectID "
    End If
    If CUSCDE = vbNullString Then
        MsgBox "There Is No Customer Control Number for this Customer", vbCritical
        Exit Sub
    End If
    SQL = Replace(SQL, "@VID", SelectCombo(cboVehicles, cboVehicles.Text))
    SQL = Replace(SQL, "@ProspectID", ProspectID)
    SQL = Replace(SQL, "@CUSCDE", N2Str2Null(CUSCDE))
    SQL = Replace(SQL, "@ProfileID", ProfileID)
    SQL = Replace(SQL, "@VehicleCode", N2Str2Null(cboVehicleCode.Text))
    SQL = Replace(SQL, "@VehicleModel", N2Str2Null(cboVehicles.Text))
    SQL = Replace(SQL, "@AcctName", N2Str2Null(AcctName))
    SQL = Replace(SQL, "@ProfileType", N2Str2Null(ProfileType))
    SQL = Replace(SQL, "@LeadSource", N2Str2Null(cboLeadSource))
    SQL = Replace(SQL, "@VehicleModel", N2Str2Null(cboVehicles))
    SQL = Replace(SQL, "@Color", N2Str2Null(cboColors))
    SQL = Replace(SQL, "@Sae", N2Str2Null(cboAttendingSE))
    SQL = Replace(SQL, "@Classification", N2Str2Null(cboClassification))
    SQL = Replace(SQL, "@Notes", N2Str2Null(txtnotes.Text))
    SQL = Replace(SQL, "@Subject", N2Str2Null(cboSubject))
    SQL = Replace(SQL, "@LogInitialInquiry", N2Str2Null(dtInitialInquiry.Value))


    gconDMIS.Execute SQL
    FillGrid 1
    picView.Visible = True: picForm.Visible = False: picSearch.Visible = False: Picture2.Visible = True
End Sub


Function GetItemData(cbo As ComboBox) As Variant
    If cbo.ListIndex = -1 Then
        GetItemData = "NULL"
    Else
        GetItemData = cbo.ItemData(cbo.ListIndex)

    End If
End Function



Private Sub frmxCustomer_ChangedData(X As Boolean)
    '''

    
    txtFilterProfile_Change
    
    LabelIt
End Sub

Private Sub frmxProfile_ChangedData(X As Boolean)

    txtFilterProfile_Change
    LabelIt
End Sub

Private Sub ImgSearchProspect_Click()
    txtFilterProfile_Change
End Sub

Private Sub lvGridCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        If lvGridCustomer.SelectedRows.Row(0).Index = 0 Then
            txtFilterProfile.SetFocus
        End If
    ElseIf KeyCode = 13 Then
        cmdSelect_Click
    ElseIf KeyCode = vbKeyEscape Then
        txtFilterProfile.SetFocus
    End If
End Sub

Private Sub lvGridCustomer_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record(6).Value = "CP" Or Row.Record(6).Value = "CC" Then
            Set frmxCustomer = New frmALLCustomer
            frmxCustomer.EditCustomer (Row.Record(7).Value)
            frmxCustomer.Show
            Set frmxCustomer = Nothing
    Else
        Set frmxProfile = New frmCRIS_EntryProfile
            frmxProfile.EditProfile ProfileID, AcctName
            frmxProfile.Show
            Set frmxProfile = Nothing
            'frmCRIS_EntryProfile.EditProfile ProfileId, AcctName
            'frmCRIS_EntryProfile.Show
    End If
End Sub

Private Sub lvGridCustomer_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    PopupMenu mnuContextTRIAD
End Sub

Private Sub lvGridCustomer_SelectionChanged()
    ProfileID = lvGridCustomer.SelectedRows.Row(0).Record(0).Value
    AcctName = lvGridCustomer.SelectedRows.Row(0).Record(1).Value
    ProfileType = lvGridCustomer.SelectedRows.Row(0).Record(6).Value
    ProspectID = 0
    CUSCDE = lvGridCustomer.SelectedRows.Row(0).Record(7).Value
    LabelIt
End Sub

'Private Sub lvGridProspect_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
''    If Row Is Nothing Then: Exit Sub
''
''    If IsNull(Row.Record(0).Record(13).Value) = False Then
''        Metrics.ForeColor = vbRed
''        Metrics.Font.Bold = True
''        'Metrics.Text.
''    Else
''
''    End If
'End Sub

Private Sub lvGridProspect_SelectionChanged()
    If lvGridProspect.Records Is Nothing Then: Exit Sub
    Dim Qstr                                 As String
    With lvGridProspect.SelectedRows.Row(0)
        ProspectID = .Record(0).Value
        ProfileType = .Record(11).Value
        ProfileID = .Record(12).Value
        
    End With
    SetProspect
    LabelIt
End Sub


Private Function SetProspect() As Boolean
    Dim temprs                               As ADODB.Recordset
    Dim Qstr                                 As String



    Set temprs = gconDMIS.Execute("Select CUSCDE, ProspectID,VehicleModel,Color, " & _
                                 "LeadSource as LeadSource, " & _
                                 "Classification as Classification, " & _
                                 "Subject as Subject, " & _
                                 "LogInitialInquiry,LogQuote,LogEmail,LogAppointment,LogTestDrive,LogCall,LogJournal,LogLetter " & _
                                 "from CRIS_PRospects Where PRospectid= " & ProspectID)

    If Not (temprs.BOF Or temprs.EOF) Then
        CUSCDE = Null2String(temprs("CUSCDE"))
        lblProspectVehicle = Null2String(temprs("VehicleModel"))
        lblProspectColor = Null2String(temprs("color"))
        lblProspectInquiry = Null2String(temprs("LogInitialInquiry"))
        lblProspectLeadSource = Null2String(temprs("LeadSource"))
        lblProspectSubject = Null2String(temprs("Subject"))
        lblProspectClassification = Null2String(temprs("Classification"))
        lblProspectStatus = vbNullString
        Qstr = vbNullString
        If IsNull(temprs!LogQuote) = False Then
            Qstr = Chr(149) & "Quotation Sent (" & temprs!LogQuote & ")" & vbCrLf
        Else

        End If

        If IsNull(temprs!LogEmail) = False Then
            Qstr = Qstr & Chr(149) & "Email Sent (" & temprs!LogEmail & ")" & vbCrLf
        Else
            Qstr = Qstr & Chr(149) & "No Email Sent" & vbCrLf
        End If

        If IsNull(temprs!LogAppointment) = False Then
            Qstr = Qstr & Chr(149) & "Appointment Made (" & temprs!LogAppointment & ")" & vbCrLf
        Else
            Qstr = Qstr & Chr(149) & "No Appointment Made " & vbCrLf
        End If

        If IsNull(temprs!LogTestDrive) = False Then
            Qstr = Qstr & Chr(149) & "Test Drive Scheduled (" & temprs!LogTestDrive & ")" & vbCrLf
        Else
            Qstr = Qstr & Chr(149) & "No Test Drive Scheduled" & vbCrLf
        End If

        If IsNull(temprs!LogCall) = False Then
            Qstr = Qstr & Chr(149) & "Calls Made (" & temprs!LogCall & ")" & vbCrLf
        Else
            Qstr = Qstr & Chr(149) & "No Calls Made" & vbCrLf
        End If

        If IsNull(temprs!LogJournal) = False Then
            Qstr = Qstr & Chr(149) & "Journals Sent:(" & temprs!LogJournal & ")" & vbCrLf
        Else
            Qstr = Qstr & Chr(149) & "No Journals Sent/Recieved " & vbCrLf
        End If

        If IsNull(temprs!LogLetter) = False Then
            Qstr = Qstr & Chr(149) & "Letter Sent (" & temprs!LogLetter & ")" & vbCrLf
        Else
            Qstr = Qstr & Chr(149) & "No Letter Recieved Or Dispatched " & vbCrLf
        End If

        lblProspectStatus = Qstr

    End If
End Function


Sub LabelIt()
    Dim temprs                               As ADODB.Recordset
    Set temprs = gconDMIS.Execute("select * from   CRIS_vW_AllProfile where Profileid=" & ProfileID & " and ProfileTYpe =" & N2Str2Null(ProfileType))

    If Not (temprs.EOF Or temprs.BOF) Then

        lblCustomerName.caption = Null2String(temprs("ProfileName").Value)
        AcctName = Null2String(temprs("AcctName").Value)
        lblAccountName.caption = Null2String(temprs("AcctName").Value)
        lblAddress.caption = Null2String(temprs("Address").Value)
        lblContactNo.caption = Null2String(temprs("Phone").Value)
        lblEmail.caption = Null2String(temprs("Email").Value)
    End If
    Set temprs = Nothing

End Sub
Private Sub mnuEditSelected_Click()
    '    If lvGridCustomer.Records(0).Record(2).Value = "CP" Or lvGridCustomer.Records(0).Record(2).Value = "CC" Then
    '        frmALLCustomer.Show vbModal
    '    End If
End Sub

Private Sub mnuRefresh_Click()
    On Error GoTo adder
    rsALLUnionCust.Requery
    Exit Sub
adder:
    Err.Clear

End Sub




Private Sub Option1_Click()

End Sub

Private Sub optShowOnly_Click(Index As Integer)
    If optShowOnly(Index).Tag = "Active" Then
        FillGrid 1
    ElseIf optShowOnly(Index).Tag = "InActive" Then
        FillGrid 2
    ElseIf optShowOnly(Index).Tag = "Closed" Then
        FillGrid 4
    Else
        FillGrid 3
    End If
End Sub

Private Sub txtFilterProfile_Change()
    Dim temprs                               As ADODB.Recordset
    ProfileID = 0
    AcctName = ""
    ProfileType = ""
    
    Set temprs = gconDMIS.Execute("Select TOP 20 * from CRIS_vW_AllProfile where AcctName like '" & txtFilterProfile.Text & "%' OR ProfileName Like '" & txtFilterProfile.Text & "%'")
    
    flex_FillReportView temprs, lvGridCustomer, False
End Sub

Private Sub txtFilterProfile_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        lvGridCustomer.SetFocus
    End If
End Sub




Private Sub txtFilterProspect_Change()
    lvGridProspect.FilterText = txtFilterProspect.Text
    lvGridProspect.Populate

    lvGridProspect.Columns(2).FooterText = txtFilterProspect.Text

End Sub
