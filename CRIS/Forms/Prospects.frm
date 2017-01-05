VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCRIS_Prospects 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14025
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
   ScaleHeight     =   8235
   ScaleWidth      =   14025
   StartUpPosition =   1  'CenterOwner
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
      Left            =   240
      MaskColor       =   &H0000FFFF&
      Picture         =   "Prospects.frx":0000
      TabIndex        =   106
      Top             =   5040
      Width           =   2130
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Reports"
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
      Left            =   210
      MaskColor       =   &H0000FFFF&
      Picture         =   "Prospects.frx":01CA
      TabIndex        =   105
      Top             =   6060
      Width           =   2130
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Inquiry"
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
      Left            =   210
      MaskColor       =   &H0000FFFF&
      Picture         =   "Prospects.frx":0394
      TabIndex        =   104
      Top             =   5580
      Width           =   2130
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Log Journal"
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
      Left            =   180
      MaskColor       =   &H0000FFFF&
      Picture         =   "Prospects.frx":055E
      TabIndex        =   102
      Top             =   3060
      Width           =   2130
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
      Height          =   420
      Left            =   180
      MaskColor       =   &H0000FFFF&
      Picture         =   "Prospects.frx":0728
      TabIndex        =   101
      Top             =   2580
      Width           =   2130
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Log Letter"
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
      Left            =   180
      MaskColor       =   &H0000FFFF&
      Picture         =   "Prospects.frx":08F2
      TabIndex        =   100
      Top             =   2100
      Width           =   2130
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
      Height          =   420
      Left            =   180
      MaskColor       =   &H0000FFFF&
      Picture         =   "Prospects.frx":0ABC
      TabIndex        =   99
      Top             =   1620
      Width           =   2130
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
      Left            =   180
      MaskColor       =   &H0000FFFF&
      Picture         =   "Prospects.frx":0C86
      TabIndex        =   98
      Top             =   1140
      Width           =   2130
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
      Left            =   180
      MaskColor       =   &H0000FFFF&
      Picture         =   "Prospects.frx":0E50
      TabIndex        =   97
      Top             =   660
      Width           =   2130
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
      Left            =   180
      MaskColor       =   &H0000FFFF&
      Picture         =   "Prospects.frx":101A
      TabIndex        =   96
      Top             =   3540
      Width           =   2130
   End
   Begin VB.CommandButton Command5 
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
      Left            =   180
      MaskColor       =   &H0000FFFF&
      Picture         =   "Prospects.frx":11E4
      TabIndex        =   82
      Top             =   180
      Width           =   2130
   End
   Begin VB.PictureBox picTestDrive 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   2910
      ScaleHeight     =   4845
      ScaleWidth      =   4905
      TabIndex        =   83
      Top             =   1770
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtBody 
         Height          =   2625
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   86
         Top             =   1500
         Width           =   4785
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   345
         Left            =   3240
         TabIndex        =   85
         Top             =   4380
         Width           =   765
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   4050
         TabIndex        =   84
         Top             =   4380
         Width           =   765
      End
      Begin MSComCtl2.DTPicker dtStartTime 
         Height          =   360
         Left            =   1860
         TabIndex        =   87
         Top             =   255
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54919170
         CurrentDate     =   39084
      End
      Begin MSComCtl2.DTPicker dtEndTime 
         Height          =   360
         Left            =   1860
         TabIndex        =   88
         Top             =   855
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54919170
         CurrentDate     =   39084
      End
      Begin MSComCtl2.DTPicker dtStartDate 
         Height          =   360
         Left            =   60
         TabIndex        =   89
         Top             =   270
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54919169
         CurrentDate     =   39084
      End
      Begin MSComCtl2.DTPicker dtEndDate 
         Height          =   360
         Left            =   60
         TabIndex        =   90
         Top             =   870
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54919169
         CurrentDate     =   39084
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   60
         TabIndex        =   91
         Top             =   4380
         Width           =   2205
         _ExtentX        =   3889
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
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   54919169
         CurrentDate     =   39084
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
         Index           =   16
         Left            =   60
         TabIndex        =   95
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "End Date /Time"
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
         Index           =   15
         Left            =   60
         TabIndex        =   94
         Top             =   645
         Width           =   1305
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Start Date/Time"
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
         Index           =   14
         Left            =   60
         TabIndex        =   93
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Next Visit Date :"
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
         Left            =   60
         TabIndex        =   92
         Top             =   4140
         Width           =   1575
      End
   End
   Begin VB.PictureBox picNextVisit 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDDADC&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   6120
      ScaleHeight     =   3705
      ScaleWidth      =   4545
      TabIndex        =   31
      Top             =   2640
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton Command3 
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
         Height          =   360
         Left            =   2730
         MaskColor       =   &H0000FFFF&
         Picture         =   "Prospects.frx":13AE
         TabIndex        =   38
         Top             =   3240
         Width           =   840
      End
      Begin VB.CommandButton Command2 
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
         Height          =   360
         Left            =   3600
         MaskColor       =   &H0000FFFF&
         Picture         =   "Prospects.frx":1578
         TabIndex        =   37
         Top             =   3240
         Width           =   840
      End
      Begin VB.TextBox Text1 
         Height          =   1875
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   1290
         Width           =   4245
      End
      Begin MSComCtl2.DTPicker dtNextVisit 
         Height          =   360
         Left            =   180
         TabIndex        =   33
         Top             =   660
         Width           =   4245
         _ExtentX        =   7488
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
         Format          =   54919169
         CurrentDate     =   39084
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Next Visit Date :"
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
         Index           =   12
         Left            =   180
         TabIndex        =   36
         Top             =   420
         Width           =   1575
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   0
         Width           =   4755
         _Version        =   655364
         _ExtentX        =   8387
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Next Visit"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Notes and Reminders for Next Visit"
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
         Index           =   4
         Left            =   180
         TabIndex        =   34
         Top             =   1050
         Width           =   4245
      End
   End
   Begin VB.PictureBox picAppointment 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   6120
      ScaleHeight     =   7545
      ScaleWidth      =   7860
      TabIndex        =   1
      Top             =   480
      Width           =   7890
      Begin VB.CommandButton Command4 
         Caption         =   "Possible Next Visit"
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
         Left            =   2910
         MaskColor       =   &H0000FFFF&
         Picture         =   "Prospects.frx":1742
         TabIndex        =   73
         Top             =   6420
         Width           =   1770
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
         Height          =   3870
         Left            =   4950
         ScaleHeight     =   3840
         ScaleWidth      =   2775
         TabIndex        =   39
         Top             =   3480
         Width           =   2805
         Begin VB.Label lblMake 
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
            Height          =   270
            Left            =   825
            TabIndex        =   60
            Top             =   1830
            Width           =   1965
         End
         Begin VB.Label lblModel 
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
            Height          =   270
            Left            =   825
            TabIndex        =   59
            Top             =   1545
            Width           =   1965
         End
         Begin VB.Label lblYear 
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
            Height          =   270
            Left            =   825
            TabIndex        =   52
            Top             =   2115
            Width           =   1965
         End
         Begin VB.Label lblColor 
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
            Height          =   270
            Left            =   825
            TabIndex        =   50
            Top             =   2400
            Width           =   1965
         End
         Begin VB.Label lblSerialNo 
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
            Height          =   270
            Left            =   825
            TabIndex        =   48
            Top             =   2685
            Width           =   1965
         End
         Begin VB.Label lblVin 
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
            Height          =   270
            Left            =   825
            TabIndex        =   46
            Top             =   2970
            Width           =   1965
         End
         Begin VB.Label lblClass 
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
            Height          =   270
            Left            =   825
            TabIndex        =   43
            Top             =   3255
            Width           =   1965
         End
         Begin VB.Label lblStatus 
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
            Height          =   270
            Left            =   825
            TabIndex        =   42
            Top             =   3540
            Width           =   1965
         End
         Begin VB.Label lblDescript 
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
            Height          =   690
            Left            =   15
            TabIndex        =   58
            Top             =   840
            Width           =   2760
         End
         Begin VB.Label lblCode 
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
            Height          =   270
            Left            =   975
            TabIndex        =   57
            Top             =   315
            Width           =   1785
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00C2FAE2&
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
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   3
            Left            =   15
            TabIndex        =   56
            Top             =   1830
            Width           =   810
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00C2FAE2&
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
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   2
            Left            =   15
            TabIndex        =   55
            Top             =   1545
            Width           =   810
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00C2FAE2&
            Caption         =   "Description"
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
            TabIndex        =   54
            Top             =   600
            Width           =   2790
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00C2FAE2&
            Caption         =   "Code"
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
            TabIndex        =   53
            Top             =   315
            Width           =   960
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00C2FAE2&
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
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   4
            Left            =   15
            TabIndex        =   51
            Top             =   2115
            Width           =   810
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00C2FAE2&
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
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   5
            Left            =   15
            TabIndex        =   49
            Top             =   2400
            Width           =   810
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00C2FAE2&
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
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   6
            Left            =   15
            TabIndex        =   47
            Top             =   2685
            Width           =   810
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00C2FAE2&
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
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   7
            Left            =   15
            TabIndex        =   45
            Top             =   2970
            Width           =   810
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00C2FAE2&
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
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   8
            Left            =   15
            TabIndex        =   44
            Top             =   3255
            Width           =   810
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00C2FAE2&
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
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   9
            Left            =   15
            TabIndex        =   41
            Top             =   3540
            Width           =   810
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   0
            Width           =   4335
            _Version        =   655364
            _ExtentX        =   7646
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Vehicles Information"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
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
         Left            =   90
         TabIndex        =   28
         Top             =   2250
         Width           =   4800
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
         Left            =   90
         TabIndex        =   25
         Top             =   2850
         Width           =   4770
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
         Height          =   420
         Left            =   3750
         MaskColor       =   &H0000FFFF&
         Picture         =   "Prospects.frx":190C
         TabIndex        =   22
         Top             =   6930
         Width           =   960
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
         Left            =   2730
         MaskColor       =   &H0000FFFF&
         Picture         =   "Prospects.frx":1AD6
         TabIndex        =   21
         Top             =   6930
         Width           =   960
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
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   4170
         Width           =   4815
      End
      Begin VB.ComboBox cboClassification 
         Height          =   330
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3510
         Width           =   4785
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
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   285
         Width           =   4800
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
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1605
         Width           =   4800
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
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   930
         Width           =   2145
      End
      Begin MSComCtl2.DTPicker dtInitialInquiry 
         Height          =   360
         Left            =   2460
         TabIndex        =   24
         Top             =   900
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   54919169
         CurrentDate     =   39084
      End
      Begin VB.PictureBox picCustomerInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   3765
         Left            =   4950
         ScaleHeight     =   3735
         ScaleWidth      =   2775
         TabIndex        =   61
         Top             =   0
         Width           =   2805
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00C2FAE2&
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
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   2
            Left            =   15
            TabIndex        =   72
            Top             =   1470
            Width           =   4380
         End
         Begin VB.Label lblCustomerName 
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   15
            TabIndex        =   71
            Top             =   615
            Width           =   2760
         End
         Begin VB.Label lblAccountName 
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   15
            TabIndex        =   70
            Top             =   1185
            Width           =   2760
         End
         Begin VB.Label lblAddress 
            BackColor       =   &H00FFFFFF&
            Height          =   690
            Left            =   15
            TabIndex        =   69
            Top             =   1740
            Width           =   2760
         End
         Begin VB.Label lblContactNo 
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   15
            TabIndex        =   68
            Top             =   2790
            Width           =   2760
         End
         Begin VB.Label lblEmail 
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   15
            TabIndex        =   67
            Top             =   3305
            Width           =   2760
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00C2FAE2&
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
            ForeColor       =   &H00400000&
            Height          =   270
            Index           =   4
            Left            =   -15
            TabIndex        =   66
            Top             =   3015
            Width           =   2790
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00C2FAE2&
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
            ForeColor       =   &H00400000&
            Height          =   270
            Index           =   3
            Left            =   15
            TabIndex        =   65
            Top             =   2445
            Width           =   2790
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00C2FAE2&
            Caption         =   "Account Name"
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
            Height          =   270
            Index           =   1
            Left            =   15
            TabIndex        =   64
            Top             =   900
            Width           =   2760
         End
         Begin VB.Label lblCustDetails 
            BackColor       =   &H00C2FAE2&
            Caption         =   "Customer Name"
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
            Height          =   270
            Index           =   0
            Left            =   15
            TabIndex        =   63
            Top             =   330
            Width           =   2760
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   315
            Index           =   3
            Left            =   0
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   0
            Width           =   2895
            _Version        =   655364
            _ExtentX        =   5106
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Customer Information"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
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
         Left            =   90
         TabIndex        =   29
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Logged"
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
         Left            =   2430
         TabIndex        =   27
         Top             =   660
         Width           =   1065
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
         Left            =   90
         TabIndex        =   26
         Top             =   2640
         Width           =   645
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
         Left            =   90
         TabIndex        =   11
         Top             =   3270
         Width           =   1110
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
         Left            =   90
         TabIndex        =   10
         Top             =   3930
         Width           =   510
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
         Left            =   90
         TabIndex        =   9
         Top             =   45
         Width           =   1680
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
         Left            =   90
         TabIndex        =   8
         Top             =   1335
         Width           =   4125
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
         Left            =   90
         TabIndex        =   7
         Top             =   690
         Width           =   1440
      End
   End
   Begin VB.PictureBox picProspectList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7590
      Left            =   2400
      ScaleHeight     =   7560
      ScaleWidth      =   3660
      TabIndex        =   74
      Top             =   480
      Width           =   3690
      Begin XtremeReportControl.ReportControl lvGridProspect 
         Height          =   3360
         Left            =   60
         TabIndex        =   76
         Top             =   990
         Width           =   3540
         _Version        =   655364
         _ExtentX        =   6244
         _ExtentY        =   5927
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         AllowColumnResize=   0   'False
      End
      Begin VB.PictureBox Picture4444 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   3150
         Left            =   60
         ScaleHeight     =   3120
         ScaleWidth      =   3495
         TabIndex        =   79
         Top             =   4380
         Width           =   3525
         Begin VB.Label lblStatusProspect 
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
            Height          =   2760
            Left            =   30
            TabIndex        =   81
            Top             =   330
            Width           =   3435
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   315
            Index           =   5
            Left            =   0
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   0
            Width           =   4335
            _Version        =   655364
            _ExtentX        =   7646
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Prospect Status Information"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            VisualTheme     =   0
            ForeColor       =   64
         End
      End
      Begin VB.TextBox txtFilterProspect 
         Height          =   345
         Left            =   60
         TabIndex        =   75
         Top             =   570
         Width           =   3510
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   4
         Left            =   0
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   0
         Width           =   3810
         _Version        =   655364
         _ExtentX        =   6720
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Prospect List"
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
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter (PRESS F8 TO REMOVE FILTER)"
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
         Left            =   45
         TabIndex        =   77
         Top             =   330
         Width           =   3000
      End
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7590
      Left            =   2430
      ScaleHeight     =   7560
      ScaleWidth      =   8640
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   8670
      Begin XtremeReportControl.ReportControl lvGridCustomer 
         Height          =   5610
         Left            =   60
         TabIndex        =   16
         Top             =   1260
         Width           =   8580
         _Version        =   655364
         _ExtentX        =   15134
         _ExtentY        =   9895
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         AllowColumnResize=   0   'False
         ShowFooter      =   -1  'True
      End
      Begin VB.CommandButton cmdaddnewcustomer 
         Caption         =   "Add New Profile"
         Height          =   345
         Left            =   5100
         TabIndex        =   23
         Top             =   870
         Width           =   1395
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Height          =   390
         Left            =   6420
         TabIndex        =   20
         Top             =   7065
         Width           =   1020
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   390
         Left            =   7560
         TabIndex        =   19
         Top             =   7065
         Width           =   1020
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Existing Prospects"
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
         Index           =   1
         Left            =   3360
         TabIndex        =   15
         Tag             =   "PP"
         Top             =   390
         Width           =   2040
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Existing Customer"
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
         Index           =   0
         Left            =   1470
         TabIndex        =   14
         Tag             =   "CP"
         Top             =   360
         Value           =   -1  'True
         Width           =   1845
      End
      Begin VB.TextBox txtFilterProfile 
         Height          =   345
         Left            =   75
         TabIndex        =   13
         Top             =   870
         Width           =   4980
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   2
         Left            =   0
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   0
         Width           =   11550
         _Version        =   655364
         _ExtentX        =   20373
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Search For Profile"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         ForeColor       =   64
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Type"
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
         Index           =   9
         Left            =   45
         TabIndex        =   18
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter"
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
         Index           =   10
         Left            =   45
         TabIndex        =   17
         Top             =   630
         Width           =   435
      End
   End
   Begin VB.Label lblCurrentProspect 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Current Prospect"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   330
      Left            =   9825
      TabIndex        =   103
      Top             =   15
      Width           =   4065
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
      Height          =   360
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _Version        =   655364
      _ExtentX        =   20558
      _ExtentY        =   635
      _StockProps     =   14
      Caption         =   " Prospect Window"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
Attribute VB_Name = "frmCRIS_Prospects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProfileId                                As Long
Dim ProfileType                              As String
Dim ProspectID                               As Long
Dim AcctName                                 As String
Event AddNewProspect(xProfileID As Long, xProfileType As String, xProspectID As Long, xWithReminder As Boolean)
Friend Sub AddNewProspect()
    ProfileId = 0
    ProfileType = vbNullString
    AcctName = vbNullString
    ProspectID = 0
    Call optSelect_Click(0)
    cmdCancelLine_Click
End Sub

Sub CenterPicture(picx As PictureBox)
    picx.Left = (Me.ScaleWidth - picx.Width) / 2
    picx.Top = (Me.ScaleHeight - picx.Height) / 2
End Sub



Private Sub cboVehicles_Click()
    If cboVehicles.ListIndex = -1 Then: Exit Sub
    Dim temprs                               As ADODB.Recordset
    Set temprs = oConSQL.Execute("SELECT * From ALL_MODEL WHERE ID=" & cboVehicles.ItemData(cboVehicles.ListIndex))
    If Not (temprs.EOF Or temprs.BOF) Then

        lblCode = Null2String(temprs!code)
        lblDescript = Null2String(temprs!descript)
        lblModel = Null2String(temprs!Model)
        lblMake = Null2String(temprs!Make)
        lblYear = Null2String(temprs!yeer)
        '        lblColor = Null2String(temprs!Color)
        '        lblSerialNo = Null2String(temprs!serialno)
        '        lblVin = Null2String(temprs!vinnumber)
        '        lblClass = Null2String(temprs!Class)
        '        lblStatus = Null2String(temprs!source)

    End If
    Set temprs = Nothing
End Sub

Private Sub cmdaddnewcustomer_Click()
    frmCRIS_EntryProfile.Show
End Sub

Private Sub cmdCancel_Click()
'    Unload Me
End Sub

Private Sub cmdCancelLine_Click()
'    picSearch.Visible = True
'    picAppointment.Visible = False
End Sub

Private Sub cmdOk_Click()
    Dim SAE                                  As String
    Dim VehicleID                            As String
    Dim Color                                As Long
    Dim PossibleNextVisit                    As String
    Dim temprs                               As ADODB.Recordset
    VehicleID = GetItemData(cboVehicles)
    SAE = GetItemData(cboAttendingSE)
    Color = GetItemData(cboColors)

    Dim SQL                                  As String

    If ProspectID > 0 Then
        SQL = " UPDATE CRIS_Prospects " _
            & " SET VehicleID=@VID, ProfileID=@ProfileID, Classification=@Classification, ProfileType=@ProfileType, AcctName=@AcctName, LeadSource=@LeadSource, VehicleModel=@VehicleModel, Color=@Color, SAE=@SAE, Notes=@Notes, Subject=@Subject, LogInitialInquiry=@LogInitialInquiry" _
            & " Where ProspectID=@ProspectID"
    Else
        SQL = "INSERT INTO CRIS_Prospects(VehicleID, Classification, ProfileType,ProfileID, AcctName, LeadSource, VehicleModel, Color, SAE, Notes, Subject, LogInitialInquiry)" _
            & " VALUES(@VID, @Classification, @ProfileType, @ProfileID, @AcctName, @LeadSource, @VehicleModel, @Color, @SAE, @Notes, @Subject, @LogInitialInquiry )" & vbCrLf & "SELECT @@IDENTITY"

    End If

    SQL = Replace(SQL, "@ProspectID", ProspectID)
    SQL = Replace(SQL, "@ProfileID", ProfileId)
    SQL = Replace(SQL, "@ProfileType", N2Str2Null(ProfileType))
    SQL = Replace(SQL, "@AcctName", N2Str2Null(AcctName))
    SQL = Replace(SQL, "@LeadSource", GetItemData(cboLeadSource))
    SQL = Replace(SQL, "@VehicleModel", N2Str2Null(cboVehicles.Text))
    SQL = Replace(SQL, "@Color", N2Str2Null(cboColors.Text))
    SQL = Replace(SQL, "@SAE", GetItemData(cboAttendingSE))
    SQL = Replace(SQL, "@Notes", N2Str2Null(txtnotes.Text))
    SQL = Replace(SQL, "@Subject", GetItemData(cboSubject))
    SQL = Replace(SQL, "@Classification", GetItemData(cboClassification))
    SQL = Replace(SQL, "@VID", GetItemData(cboVehicles))
    SQL = Replace(SQL, "@LogInitialInquiry", N2Str2Null(dtInitialInquiry.Value))
    Set temprs = oConSQL.Execute(SQL)
    If ProspectID <= 0 Then
        MessagePop RecSaveOk, "Record Added ", "New Prospect Added Sucessfully ", 500, 1
    Else
        MessagePop RecSaveOk, "RecordSaved", "Prospect Information Sucessfully Updated", 500, 1
    End If
    Set temprs = temprs.NextRecordset
    If Not temprs Is Nothing Then
        ProspectID = temprs.Collect(0)
  
    End If


    Set temprs = Nothing
    FillGrid
picProspectList.Enabled = True


End Sub

Private Sub cmdSelect_Click()

ShowHide picSearch.hwnd, False
    'picSearch.Visible = False
    'picAppointment.Visible = True
    lvGridCustomer_SelectionChanged
    cboVehicles.SetFocus
End Sub

Private Sub Command1_Click()
    ShowHide picSearch.hwnd, False
End Sub

Private Sub Command10_Click()
    frmCRIS_LogCall.Show
End Sub

Private Sub Command12_Click()
frmCRIS_LogEmail.Show
End Sub

Private Sub Command13_Click()
frmCRIS_LogJournal.Show
End Sub

Private Sub Command16_Click()
    frmCRIS_AOR.Show
End Sub

Private Sub Command2_Click()
    ShowHide picNextVisit.hwnd, False
End Sub

Private Sub Command4_Click()
    ShowHide picNextVisit.hwnd, True
End Sub

Private Sub Command5_Click()

ShowHide picSearch.hwnd, True
    AddNewProspect
End Sub

Private Sub Command7_Click()
    frmCRIS_EntrySalesAppointment.Show
End Sub

Private Sub Command8_Click()
    
    frmCRIS_EntryTestDriveAppointment.Show
End Sub

Private Sub Command9_Click()
    frmCRIS_EntryQuotation.Show
End Sub

Private Sub Form_Load()
    InitVars
    CenterPicture picNextVisit
    CenterPicture picTestDrive
  '  CenterPicture picSearch
End Sub

Function GetItemData(cbo As ComboBox) As Variant
    If cbo.ListIndex = -1 Then
        GetItemData = -1
    Else
        GetItemData = cbo.ItemData(cbo.ListIndex)

    End If
End Function

Private Function GetProspect() As Boolean

End Function

Private Sub InitVars()
    Dim SQL                                  As String
    Dim temprs                               As ADODB.Recordset
    Call FillCombo("Select DataID as ID, " _
                 & " MasterData as [Description] " _
                 & " from CRIS_vw_master_PullDown Where MasterType='Inquiry Type'", 0, 1, cboSubject)

    Call FillCombo("Select DISTINCT 1, COLOR_DESC FROM ALL_COLOR ORDER BY COLOR_DESC", 0, 1, cboColors)
    Call FillCombo("SELECT ID, DESCRIPT from ALL_MODEL", 0, 1, cboVehicles)
    Call FillCombo("SELECT ID, LastName + ', ' + FirstName + '.'+ MiddleName from HRMS_EMPINFO WHERE IS_SAE=1 ORDER BY LASTNAME", 0, 1, cboAttendingSE)
    
    Set temprs = oConSQL.Execute("Select DataID, MasterData ,MasterType from CRIS_vw_master_PullDown where MasterType IN ('Customer Classification', 'Lead Source')")
    
    While Not temprs.EOF
        If temprs.Fields("MasterType").Value = "Lead Source" Then
            cboLeadSource.AddItem temprs.Collect(1)
            cboLeadSource.ItemData(cboLeadSource.NewIndex) = temprs.Fields(0).Value
        Else
            cboClassification.AddItem temprs.Collect(1)
            cboClassification.ItemData(cboClassification.NewIndex) = temprs.Fields(0).Value

        End If
        temprs.MoveNext

    Wend
    With lvGridCustomer
        .Columns.Add 0, "ID", 0, True
        .Columns.Add 1, "Account Name", 50, True
        .Columns.Add 2, "Profile Name", 100, True
        .Columns(0).Visible = False

        .PaintManager.HorizontalGridStyle = xtpGridSmallDots  ' xtpGridNoLines
        .PaintManager.GroupRowTextBold = True                 ' = vbWhite
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True

    End With


    With lvGridProspect
        .Columns.Add 0, "ID", 0, True
        .Columns.Add 1, "Date", 50, True
        .Columns.Add 2, "AccountName", 100, True
        '  .Columns.Add 3, "Model", 100, True

        .Columns(0).Visible = False

        .PaintManager.HorizontalGridStyle = xtpGridSmallDots  ' xtpGridNoLines
        .PaintManager.GroupRowTextBold = True                 ' = vbWhite
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True

    End With

    FillGrid

End Sub
Sub FillGrid()
    Dim temprs                               As ADODB.Recordset
    Set temprs = oConSQL.Execute("SELECT  ProspectID, " & _
                               " convert(varchar, LogInitialInquiry ,101)as [Date] ,  " & _
                               " AcctName, VehicleModel, LogQuote, " & _
                               " LogEmail, LogAppointment, LogTestDrive, " & _
                               " LogCall, LogJournal, LogLetter, ProfileType , ProfileID " & _
                               " FROM CRIS_Prospects " & _
                               " Where D_S is NULL ")
    flex_FillReportView temprs, lvGridProspect, False
End Sub
Sub LabelIt()
    Dim temprs                               As ADODB.Recordset
    Set temprs = oConSQL.Execute("select * from   CRIS_vW_AllProfile where Profileid=" & ProfileId & " and ProfileTYpe =" & N2Str2Null(ProfileType))

    If Not (temprs.EOF Or temprs.BOF) Then
        lblCurrentProspect = Null2String(temprs("ProfileName").Value)
        lblCustomerName = Null2String(temprs("ProfileName").Value)
        lblAccountName = Null2String(temprs("AcctName").Value)
        lblAddress = Null2String(temprs("Address").Value)
        lblContactNo = Null2String(temprs("Phone").Value)
        lblEmail = Null2String(temprs("Email").Value)
    End If
    Set temprs = Nothing

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
    '    If Row.Record(2).Value = "CP" Or Row.Record(2).Value = "CC" Then
    '        frmCRIS_Customer.Show vbModal
    '    End If
    cmdSelect_Click
End Sub

Private Sub lvGridCustomer_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    PopupMenu mnuContextTRIAD
End Sub

Private Sub lvGridCustomer_SelectionChanged()
    ProfileId = lvGridCustomer.SelectedRows.Row(0).Record(0).Value
    ProfileType = lvGridCustomer.SelectedRows.Row(0).Record(6).Value
    AcctName = lvGridCustomer.SelectedRows.Row(0).Record(1).Value
Debug.Print ProfileId; ProfileType; AcctName

    LabelIt

End Sub

Private Sub lvGridProspect_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If lvGridProspect.Records Is Nothing Then Exit Sub

    SetProspect CLng(Row.Record(0).Value)

    picSearch.Visible = False
    picAppointment.Visible = True
    cboVehicles.SetFocus
    picProspectList.Enabled = False

End Sub

Private Sub lvGridProspect_SelectionChanged()
    lblStatusProspect.caption = vbNullString
    If lvGridProspect.Records Is Nothing Then: Exit Sub

    Dim Qstr                                 As String

    With lvGridProspect.SelectedRows.Row(0)

        If IsNull(.Record(4).Value) = False Then: Qstr = Chr(149) & "Quotation Sent (" & .Record(4).Value & ")" & vbCrLf
        If IsNull(.Record(5).Value) = False Then: Qstr = Qstr & Chr(149) & "Email Sent (" & .Record(5).Value & ")" & vbCrLf
        If IsNull(.Record(6).Value) = False Then: Qstr = Qstr & Chr(149) & "Appointment Made (" & .Record(6).Value & ")" & vbCrLf
        If IsNull(.Record(7).Value) = False Then: Qstr = Qstr & Chr(149) & "Test Drive Scheduled (" & .Record(7).Value & ")" & vbCrLf
        If IsNull(.Record(8).Value) = False Then: Qstr = Qstr & Chr(149) & "Calls Made (" & .Record(8).Value & ")" & vbCrLf
        If IsNull(.Record(9).Value) = False Then: Qstr = Qstr & Chr(149) & "Journals Added (" & .Record(9).Value & ")" & vbCrLf
        If IsNull(.Record(10).Value) = False Then: Qstr = Qstr & Chr(149) & "Letter Sent (" & .Record(10).Value & ")" & vbCrLf

        lblStatusProspect = Qstr
        ProspectID = .Record(0).Value
        ProfileType = .Record(11).Value
        ProfileId = .Record(12).Value
        AcctName = .Record(2).Value

    
    

    SetProspect ProspectID

    picSearch.Visible = False
    picAppointment.Visible = True
    
    
    
    

    End With
End Sub

Private Sub mnuEditSelected_Click()
    '    If lvGridCustomer.Records(0).Record(2).Value = "CP" Or lvGridCustomer.Records(0).Record(2).Value = "CC" Then
    '        frmALLCustomer.Show vbModal
    '    End If
End Sub

Private Sub mnuRefresh_Click()
    ''

End Sub


Private Function SetProspect(nProspectID As Long) As Boolean
    Dim SQL                             As String
    Dim temprs                          As ADODB.Recordset

    Set temprs = oConSQL.Execute("Select * from CRIS_PROSPECTS Where ProspectID=" & nProspectID)
    If Not (temprs.BOF Or temprs.EOF) Then
        cboVehicles.Text = Null2String(temprs("VehicleModel"))
        cboColors.ListIndex = SelectCombo(cboColors, Null2String(temprs("color")), False)
        dtInitialInquiry.Value = Null2String(temprs("LogInitialInquiry"))
        cboAttendingSE.ListIndex = SelectCombo(cboAttendingSE, Null2String(temprs("SAE")), True)
        cboLeadSource.ListIndex = SelectCombo(cboLeadSource, Null2String(temprs("LeadSource")), True)
        cboSubject.ListIndex = SelectCombo(cboSubject, Null2String(temprs("Subject")), True)
        txtnotes.Text = Null2String(temprs("Notes"))
        cboClassification.ListIndex = SelectCombo(cboClassification, Null2String(temprs("Classification")), True)
    End If
    LabelIt
End Function

Sub ShowHide(hwnd As Long, State As Boolean)
    Dim cntl                                 As Control
    For Each cntl In Me.Controls
        If TypeOf cntl Is PictureBox Then
            If Not cntl.Container.hwnd = hwnd Then
                If cntl.hwnd = hwnd Then
                    cntl.Enabled = State
                    cntl.Visible = State
                    If State = True Then
                        cntl.ZOrder 0
                    Else
                        cntl.ZOrder 1
                    End If
                Else
                    cntl.Enabled = Not (State)

                End If
            End If
        End If
    Next
End Sub

Private Sub optSelect_Click(Index As Integer)
    If optSelect(Index).Value = True Then
        txtFilterProfile_Change
    End If
End Sub

Private Sub txtFilterProfile_Change()

    Dim temprs                               As ADODB.Recordset
    If optSelect(0).Value = True Then
        Set temprs = oConSQL.Execute("Select TOP 20 * from CRIS_vW_AllProfile where ProfileType NOT in ('PP','PC') and  AcctName like '%" & txtFilterProfile.Text & "%'")
    Else
        Set temprs = oConSQL.Execute("Select TOP 20  * from CRIS_vW_AllProfile where ProfileType  in ('PP','PC') and  AcctName like '%" & txtFilterProfile.Text & "%'")
        'ProfileId , AcctName, ProfileName, Address, Email, Phone, ProfileType
    End If

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
End Sub
