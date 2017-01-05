VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_EntryTestDriveAppointment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Drive"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12435
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
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   12435
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picInfoView 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7185
      ScaleWidth      =   5250
      TabIndex        =   37
      Top             =   0
      Width           =   5280
      Begin VB.PictureBox picProfileCard 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   60
         ScaleHeight     =   2385
         ScaleWidth      =   5100
         TabIndex        =   55
         Top             =   390
         Width           =   5130
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
            TabIndex        =   66
            Top             =   2115
            Width           =   1500
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
            TabIndex        =   65
            Top             =   1830
            Width           =   1500
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
            TabIndex        =   64
            Top             =   615
            Width           =   1515
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
            TabIndex        =   63
            Top             =   315
            Width           =   1515
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   62
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
            TabIndex        =   61
            Top             =   885
            Width           =   5070
         End
         Begin VB.Label lblCustomerName 
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
            Height          =   270
            Left            =   1545
            TabIndex        =   60
            Top             =   320
            Width           =   3555
         End
         Begin VB.Label lblAccountName 
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
            Height          =   270
            Left            =   1545
            TabIndex        =   59
            Top             =   600
            Width           =   3555
         End
         Begin VB.Label lblAddress 
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
            Height          =   690
            Left            =   15
            TabIndex        =   58
            Top             =   1125
            Width           =   5070
         End
         Begin VB.Label lblContactNo 
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
            Height          =   270
            Left            =   1530
            TabIndex        =   57
            Top             =   1830
            Width           =   3555
         End
         Begin VB.Label lblEmail 
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
            Height          =   270
            Left            =   1530
            TabIndex        =   56
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
         TabIndex        =   52
         Top             =   5640
         Width           =   5115
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   315
            Index           =   5
            Left            =   0
            TabIndex        =   54
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
            TabIndex        =   53
            Top             =   330
            Width           =   5025
         End
      End
      Begin VB.PictureBox picProspectCard 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   2715
         Left            =   60
         ScaleHeight     =   2685
         ScaleWidth      =   5100
         TabIndex        =   38
         Top             =   2880
         Width           =   5130
         Begin VB.Label lblProspectInquiry 
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   1530
            TabIndex        =   51
            Top             =   2115
            Width           =   3555
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
            TabIndex        =   50
            Top             =   1830
            Width           =   3555
         End
         Begin VB.Label lblProspectSubject 
            BackColor       =   &H00FFFFFF&
            Height          =   690
            Left            =   15
            TabIndex        =   49
            Top             =   1125
            Width           =   5070
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
            TabIndex        =   47
            Top             =   885
            Width           =   5100
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   315
            Index           =   3
            Left            =   0
            TabIndex        =   46
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
            TabIndex        =   45
            Top             =   315
            Width           =   1515
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
            TabIndex        =   44
            Top             =   1830
            Width           =   1500
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
            TabIndex        =   43
            Top             =   2115
            Width           =   1500
         End
         Begin VB.Label lblProspectClassification 
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   1530
            TabIndex        =   42
            Top             =   2400
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
            TabIndex        =   41
            Top             =   2400
            Width           =   1500
         End
         Begin VB.Label lblProspectColor 
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   1545
            TabIndex        =   40
            Top             =   600
            Width           =   3555
         End
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
            TabIndex        =   39
            Top             =   600
            Width           =   1515
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   2
         Left            =   -90
         TabIndex        =   67
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
   Begin VB.ComboBox cboClassification 
      Height          =   330
      Left            =   10200
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   2820
      Width           =   2220
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   390
      Left            =   11340
      MaskColor       =   &H0000FFFF&
      Picture         =   "TestDriveAppointment.frx":0000
      TabIndex        =   17
      Top             =   6690
      Width           =   1020
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
      Height          =   375
      Left            =   10260
      MaskColor       =   &H0000FFFF&
      Picture         =   "TestDriveAppointment.frx":01CA
      TabIndex        =   16
      Top             =   6690
      Width           =   1020
   End
   Begin VB.TextBox txtFeedBack 
      Height          =   1725
      Left            =   9090
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4260
      Width           =   3285
   End
   Begin VB.TextBox txtInterest 
      Height          =   1710
      Left            =   5400
      TabIndex        =   10
      Top             =   4260
      Width           =   3600
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
      ForeColor       =   &H00400000&
      Height          =   345
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2880
      Width           =   4680
   End
   Begin VB.PictureBox picVehiclesDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   2535
      Left            =   5310
      ScaleHeight     =   2505
      ScaleWidth      =   7005
      TabIndex        =   0
      Top             =   0
      Width           =   7035
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
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   3990
      End
      Begin VB.Label lblStatus 
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
         Height          =   270
         Left            =   4920
         TabIndex        =   32
         Top             =   2205
         Width           =   2055
      End
      Begin VB.Label lblCapDetails 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Source"
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
         Left            =   4005
         TabIndex        =   31
         Top             =   2205
         Width           =   900
      End
      Begin VB.Label lblClass 
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
         Height          =   270
         Left            =   900
         TabIndex        =   30
         Top             =   2205
         Width           =   3075
      End
      Begin VB.Label lblCapDetails 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Class"
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
         Left            =   -15
         TabIndex        =   29
         Top             =   2205
         Width           =   900
      End
      Begin VB.Label lblVin 
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
         Height          =   270
         Left            =   4920
         TabIndex        =   28
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblCapDetails 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Vin"
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
         Left            =   4005
         TabIndex        =   27
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label lblYear 
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
         Height          =   270
         Left            =   4920
         TabIndex        =   26
         Top             =   1350
         Width           =   2055
      End
      Begin VB.Label lblCapDetails 
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
         Index           =   4
         Left            =   4005
         TabIndex        =   25
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label lblMake 
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
         Height          =   270
         Left            =   4920
         TabIndex        =   24
         Top             =   1065
         Width           =   2055
      End
      Begin VB.Label lblModel 
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
         Height          =   270
         Left            =   4920
         TabIndex        =   23
         Top             =   780
         Width           =   2055
      End
      Begin VB.Label lblDescript 
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
         Height          =   690
         Left            =   15
         TabIndex        =   22
         Top             =   1215
         Width           =   3960
      End
      Begin VB.Label lblCode 
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
         Height          =   270
         Left            =   4920
         TabIndex        =   21
         Top             =   495
         Width           =   2055
      End
      Begin VB.Label lblColor 
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
         Height          =   270
         Left            =   4920
         TabIndex        =   20
         Top             =   1635
         Width           =   2055
      End
      Begin VB.Label lblSerialNo 
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
         Height          =   270
         Left            =   900
         TabIndex        =   19
         Top             =   1920
         Width           =   3075
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   4
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         Width           =   7080
         _Version        =   655364
         _ExtentX        =   12488
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Select Test Drive Vehicles"
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
      Begin VB.Label lblCapDetails 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Serial No"
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
         Left            =   -15
         TabIndex        =   6
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label lblCapDetails 
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
         Index           =   5
         Left            =   4005
         TabIndex        =   5
         Top             =   1635
         Width           =   900
      End
      Begin VB.Label lblCapDetails 
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
         Index           =   0
         Left            =   4005
         TabIndex        =   4
         Top             =   495
         Width           =   900
      End
      Begin VB.Label lblCapDetails 
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
         Index           =   1
         Left            =   15
         TabIndex        =   3
         Top             =   975
         Width           =   3960
      End
      Begin VB.Label lblCapDetails 
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
         Index           =   2
         Left            =   4005
         TabIndex        =   2
         Top             =   780
         Width           =   900
      End
      Begin VB.Label lblCapDetails 
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
         Index           =   3
         Left            =   4005
         TabIndex        =   1
         Top             =   1065
         Width           =   900
      End
      Begin VB.Label lblCapDetails 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Test Drive Vehicles"
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
         Left            =   0
         TabIndex        =   68
         Top             =   330
         Width           =   4905
      End
   End
   Begin VB.PictureBox picEventDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   5400
      ScaleHeight     =   765
      ScaleWidth      =   7095
      TabIndex        =   13
      Top             =   3270
      Width           =   7095
      Begin MSComCtl2.DTPicker dtStartTime 
         Height          =   360
         Left            =   3555
         TabIndex        =   14
         Top             =   210
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm:ss"
         Format          =   54853634
         CurrentDate     =   39084
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   360
         Left            =   60
         TabIndex        =   36
         Top             =   240
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54853632
         CurrentDate     =   39084
      End
      Begin MSComCtl2.DTPicker dtEndTime 
         Height          =   360
         Left            =   5310
         TabIndex        =   69
         Top             =   210
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54853634
         CurrentDate     =   39084
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Date"
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
         Left            =   45
         TabIndex        =   15
         Top             =   60
         Width           =   405
      End
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
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
      Left            =   10170
      TabIndex        =   35
      Top             =   2610
      Width           =   1110
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      Caption         =   "Feed Back"
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
      Left            =   9120
      TabIndex        =   34
      Top             =   4050
      Width           =   855
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      Caption         =   "Special Interest  Concerns"
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
      Index           =   6
      Left            =   5400
      TabIndex        =   11
      Top             =   3990
      Width           =   2220
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      Caption         =   "Attending SAE"
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
      Index           =   8
      Left            =   5400
      TabIndex        =   9
      Top             =   2640
      Width           =   1200
   End
End
Attribute VB_Name = "frmCRIS_EntryTestDriveAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProspectID                               As Long
Dim ProfileID                                As Long
Dim ProfileType                              As String
Dim AcctName                                 As String
Dim TestDriveScheduleID                      As Long


''''''CALLS
Friend Sub AddTestDriveAppointment(xProspectID As Long, xAcctName As String, xProfileType As String, xProfileID As Long)
    ProspectID = xProspectID
    AcctName = xAcctName
    ProfileType = xProfileType
    TestDriveScheduleID = 0
    ProfileID = xProfileID
    LabelIt
    SetProspect
End Sub

Friend Sub EditTestDriveAppointment(xProspectID As Long, xAcctName As String, xTestDriveScheduleID As Long)
    ProspectID = xProspectID
    AcctName = xAcctName
    TestDriveScheduleID = xTestDriveScheduleID
End Sub
'''END CALLS

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


Private Function SetProspect() As Boolean
    Dim temprs                               As ADODB.Recordset
    Dim Qstr                                 As String



    Set temprs = gconDMIS.Execute("Select  ProspectID,VehicleModel,Color, " & _
                                 "LeadSource as LeadSource, " & _
                                 "Classification, " & _
                                 "Subject, " & _
                                 "LogInitialInquiry,LogQuote,LogEmail,LogAppointment,LogTestDrive,LogCall,LogJournal,LogLetter " & _
                                 "from CRIS_PRospects Where PRospectid= " & ProspectID)

    If Not (temprs.BOF Or temprs.EOF) Then
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
Private Sub cboVehicles_Click()
    FillVehicleCard
End Sub
Sub FillVehicleCard()
    Dim temprs                               As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT * From CRIS_MRRINV WHERE ID=" & cboVehicles.ItemData(cboVehicles.ListIndex))
    If Not (temprs.EOF Or temprs.BOF) Then
        lblCode = Null2String(temprs!code)
        lblDescript = Null2String(temprs!descript)
        lblModel = Null2String(temprs!Model)
        lblMake = Null2String(temprs!Make)
        lblYear = Null2String(temprs!yearmodel)
        lblColor = Null2String(temprs!Color)
        lblSerialNo = Null2String(temprs!serialno)
        lblVin = Null2String(temprs!vinnumber)
        lblClass = Null2String(temprs!Class)
        lblStatus = Null2String(temprs!Source)
    End If
    Set temprs = Nothing
End Sub

Private Sub Combo1_Click()

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub





Private Sub Form_Load()
    InitVars
End Sub
Private Sub InitVars()
    Dim SQL                                  As String
    Dim temprs                               As ADODB.Recordset

    dtStartTime.Value = WorkTimeStart
    dtStartTime.MinDate = DateFromString("01/01/2007", "8:00:00 AM")
    dtStartTime.MaxDate = DateFromString("01/01/2007", "5:30:00 PM")
    dtEndTime.Value = DateAdd("h", 1, WorkTimeStart)
    dtEndTime.MinDate = DateFromString("01/01/2007", CStr(WorkTimeStart))
    dtEndTime.MaxDate = DateFromString("01/01/2007", CStr(WorkTimeEnd))



    Call FillCombo("Select ID, CODE  from cris_mrrInv where DateReturned is Null and code<>'NULL' order by code ", 0, 1, cboVehicles)
    Call FillCombo("SELECT ID, LastName + ', ' + FirstName + '.'+ MiddleName from HRMS_EMPINFO WHERE IS_SAE=1 ORDER BY LASTNAME", 0, 1, cboAttendingSE)
    Set temprs = gconDMIS.Execute("Select MasterData from CRIS_vw_master_PullDown where MasterType='Customer Classification'")
    While Not temprs.EOF
        cboClassification.AddItem temprs.Collect(0)
        cboClassification.ItemData(cboClassification.NewIndex) = cboClassification.ListCount
        temprs.MoveNext
    Wend

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
    If TestDriveScheduleID <= 0 Then
        SQL = "INSERT INTO CRIS_TestDriveSchedules " & _
            " (ProspectID, VehicleID, SAE, PossibleNextVisit, NextVisitNotes, Interests, FeedBack, StartDateTime,EndDateTime, Classification) " & _
            " VALUES(@ProspectID, @VehicleID, @SAE, @PossibleNextVisit, @NextVisitNotes, @Interests, @FeedBack, @StartDateTime, @EndDateTime, @Classification) " & vbCrLf & " SELECT @@IDENTITY"
    Else
        SQL = "Update CRIS_TestDriveSchedules " & _
              "SET  ProspectID=@ProspectID, VehicleID=@VehicleID, SAE=@SAE, PossibleNextVisit=@PossibleNextVisit, NextVisitNotes=@NextVisitNotes, Interests=@Interests, StartDateTime=@StartDateTime, Classification=@Classification" & _
            " WHERE ScheduleID=@ScheduleID"
    End If
    Dim t1 As String, t2                     As String
    t1 = FormatDateTime(dtDate.Value, vbShortDate) & " " & FormatDateTime(dtStartTime.Value, vbLongTime)
    t2 = FormatDateTime(dtDate.Value, vbShortDate) & " " & FormatDateTime(dtStartTime.Value, vbLongTime)

    SQL = Replace(SQL, "@ScheduleID", TestDriveScheduleID)
    SQL = Replace(SQL, "@ProspectID", ProspectID)
    SQL = Replace(SQL, "@VehicleID", GetItemData(cboVehicles))
    SQL = Replace(SQL, "@SAE", GetItemData(cboAttendingSE))
    SQL = Replace(SQL, "@PossibleNextVisit", N2Str2Null(""))
    SQL = Replace(SQL, "@NextVisitNotes", N2Str2Null(""))
    SQL = Replace(SQL, "@Interests", N2Str2Null(txtInterest))
    SQL = Replace(SQL, "@FeedBack", N2Str2Null(txtFeedBack))
    SQL = Replace(SQL, "@StartDateTime", N2Str2Null(t1))
    SQL = Replace(SQL, "@EndDateTime", N2Str2Null(t2))
    SQL = Replace(SQL, "@Classification", GetItemData(cboClassification))




    Dim temprs                               As ADODB.Recordset

    Set temprs = gconDMIS.Execute(SQL)
    gconDMIS.Execute ("update CRIS_PROSPECTs SET LogTestDrive=" & N2Str2Null(t1) & " where prospectid=" & ProspectID)

    If TestDriveScheduleID <= 0 Then
        MessagePop RecSaveOk, "Record Added ", "New Schedule Sucessfully Added", 500, 1
    Else
        MessagePop RecSaveOk, "RecordSaved", "Schedule Sucessfully Updated", 500, 1
    End If

    Set temprs = temprs.NextRecordset
    If Not temprs Is Nothing Then
        TestDriveScheduleID = temprs.Collect(0)
    End If


    Set temprs = Nothing

End Sub

Private Sub cboClassification_Click()
    Dim pLabel                               As CalendarEventLabel
    Dim nLabelID                             As Long

    nLabelID = cboClassification.ItemData(cboClassification.ListIndex)


End Sub

Private Sub chkAllDayEvent_Click()
End Sub

Private Function CheckDates() As Boolean
    CheckDates = True
    If (Not IsDateValid(cboStartDate.Text)) Then
        cboStartDate.SetFocus
        CheckDates = False
        Exit Function
    End If

    If (Not IsDateValid(cboEndDate.Text)) Then
        cboEndDate.SetFocus
        CheckDates = False
        Exit Function
    End If

End Function
Public Sub ModifyEvent(ModEvent As CalendarEvent)
    Set EventCalendar = ModEvent
    AddEvent = False
End Sub
Function GetItemData(cbo As ComboBox) As Variant
    If cbo.ListIndex = -1 Then
        GetItemData = -1
    Else
        GetItemData = cbo.ItemData(cbo.ListIndex)

    End If
End Function

Private Sub lvGridCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub lvGridCustomer_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record(2).Value = "CP" Or Row.Record(2).Value = "CC" Then
    End If
End Sub

Private Sub lvGridCustomer_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    'PopupMenu mnucon
End Sub

Private Sub lvGridCustomer_SelectionChanged()
End Sub


Private Sub optSelect_Click(Index As Integer)
    Dim temprs                               As ADODB.Recordset
End Sub


Private Sub txtFilterProfile_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
    End If
End Sub



