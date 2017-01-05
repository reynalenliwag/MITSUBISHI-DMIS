VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Begin VB.Form frmSMIS_Trans_MRR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Receiving Entry"
   ClientHeight    =   9390
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   13170
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
   Icon            =   "MRRINV.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   13170
   Begin VB.PictureBox picServiceInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5385
      Left            =   2880
      ScaleHeight     =   5355
      ScaleWidth      =   6645
      TabIndex        =   151
      Top             =   1560
      Visible         =   0   'False
      Width           =   6675
      Begin VB.TextBox txtVehDet 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   164
         Top             =   1890
         Width           =   6555
      End
      Begin VB.TextBox txtVeh_register_CustName 
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   157
         Top             =   660
         Width           =   5595
      End
      Begin VB.CommandButton cmdRegisterVehicleClose 
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
         Height          =   795
         Left            =   5880
         MouseIcon       =   "MRRINV.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "MRRINV.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   155
         ToolTipText     =   "Cancel"
         Top             =   1080
         Width           =   705
      End
      Begin VB.TextBox txtVeh_register_CustCode 
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
         Left            =   90
         MaxLength       =   6
         TabIndex        =   153
         Text            =   "123456"
         Top             =   660
         Width           =   855
      End
      Begin VB.CommandButton cmdRegisterVehicle 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   795
         Left            =   5190
         MouseIcon       =   "MRRINV.frx":0FC1
         MousePointer    =   99  'Custom
         Picture         =   "MRRINV.frx":1113
         Style           =   1  'Graphical
         TabIndex        =   156
         ToolTipText     =   "Save this Record"
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Detail Vehicle Info "
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
         Height          =   255
         Left            =   90
         TabIndex        =   165
         Top             =   1620
         Width           =   1935
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
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
         Height          =   255
         Left            =   90
         TabIndex        =   154
         Top             =   390
         Width           =   1935
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   405
         Left            =   0
         TabIndex        =   152
         Top             =   -60
         Width           =   6645
         _Version        =   655364
         _ExtentX        =   11721
         _ExtentY        =   714
         _StockProps     =   14
         Caption         =   "Update Service Vehicle Info"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picTops 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   18825
      TabIndex        =   0
      Top             =   0
      Width           =   18825
      Begin VB.PictureBox picRefHeader 
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   9780
         ScaleHeight     =   2115
         ScaleWidth      =   4845
         TabIndex        =   22
         Top             =   60
         Width           =   4845
         Begin VB.CommandButton Command5 
            Caption         =   "::"
            Height          =   345
            Left            =   2970
            TabIndex        =   30
            Top             =   1110
            Width           =   345
         End
         Begin VB.TextBox txtDRNO 
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
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   1185
            MaxLength       =   15
            TabIndex        =   24
            Top             =   0
            Width           =   2115
         End
         Begin VB.TextBox txtref_PONO 
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
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1185
            MaxLength       =   15
            TabIndex        =   26
            Top             =   375
            Width           =   2115
         End
         Begin MSMask.MaskEdBox txtDateReleased 
            Height          =   345
            Left            =   1200
            TabIndex        =   33
            ToolTipText     =   "Date Vehicles Released"
            Top             =   1500
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   -2147483633
            ForeColor       =   7347754
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd-mmm-yy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDateReceived 
            Height          =   345
            Left            =   1200
            TabIndex        =   31
            Tag             =   "@R"
            ToolTipText     =   "Date Vehicles Received (Recieved Date)"
            Top             =   1125
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd-mmm-yy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPullOutDate 
            Height          =   345
            Left            =   1200
            TabIndex        =   27
            ToolTipText     =   "Date of Pull Out ( Pull Out Date)"
            Top             =   750
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd-mmm-yy"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton Command4 
            Caption         =   "::"
            Height          =   345
            Left            =   2970
            TabIndex        =   29
            Top             =   750
            Width           =   345
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Ref: INV#."
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
            Left            =   165
            TabIndex        =   25
            Top             =   420
            Width           =   810
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Ref: DR NO."
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
            Left            =   165
            TabIndex        =   23
            Top             =   45
            Width           =   960
         End
         Begin VB.Label Label39 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Pull Out"
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
            Height          =   255
            Left            =   165
            TabIndex        =   28
            Top             =   810
            Width           =   1185
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Released"
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
            Height          =   255
            Left            =   165
            TabIndex        =   34
            Top             =   1545
            Width           =   1185
         End
         Begin VB.Label Label38 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Received"
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
            Height          =   255
            Left            =   165
            TabIndex        =   32
            Top             =   1215
            Width           =   1185
         End
      End
      Begin VB.PictureBox picModelDetails 
         BorderStyle     =   0  'None
         Height          =   2040
         Left            =   30
         ScaleHeight     =   2040
         ScaleWidth      =   6945
         TabIndex        =   1
         Top             =   -30
         Width           =   6945
         Begin VB.CommandButton cmdAddFromPO 
            Height          =   345
            Left            =   6555
            Picture         =   "MRRINV.frx":171A
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   60
            Width           =   345
         End
         Begin VB.TextBox txtPO 
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
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   45
            Width           =   2175
         End
         Begin VB.ComboBox cboTransmission 
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
            ItemData        =   "MRRINV.frx":18E4
            Left            =   5280
            List            =   "MRRINV.frx":18EE
            TabIndex        =   17
            Top             =   1560
            Width           =   1665
         End
         Begin VB.TextBox txtModelCode 
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
            Left            =   1065
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1203
            Width           =   750
         End
         Begin VB.TextBox txtYeer 
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
            Height          =   330
            Left            =   5295
            MaxLength       =   4
            TabIndex        =   15
            Top             =   1185
            Width           =   1620
         End
         Begin VB.TextBox txtMake 
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
            Height          =   330
            Left            =   1065
            TabIndex        =   8
            Top             =   446
            Width           =   5835
         End
         Begin VB.TextBox txtCode 
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
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   1065
            TabIndex        =   4
            Top             =   45
            Width           =   1515
         End
         Begin VB.ComboBox cboClass 
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
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1590
            Width           =   2925
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
            Left            =   1065
            TabIndex        =   10
            Text            =   "txtDescript"
            Top             =   817
            Width           =   5850
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
            Left            =   1860
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1215
            Width           =   2130
         End
         Begin VB.Label Label32 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   2
            Left            =   330
            TabIndex        =   144
            Top             =   1290
            Width           =   150
         End
         Begin VB.Label Label32 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   1
            Left            =   30
            TabIndex        =   143
            Top             =   660
            Width           =   150
         End
         Begin VB.Label Label32 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   142
            Top             =   480
            Width           =   150
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "PO NO"
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
            Left            =   3705
            TabIndex        =   2
            Top             =   90
            Width           =   555
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Transmission"
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
            Left            =   4065
            TabIndex        =   19
            Top             =   1650
            Width           =   1170
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Class"
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
            Left            =   540
            TabIndex        =   16
            Top             =   1635
            Width           =   480
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
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
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   4785
            TabIndex        =   14
            Top             =   1260
            Width           =   390
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
            Left            =   510
            TabIndex        =   11
            Top             =   1260
            Width           =   510
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
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
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   555
            TabIndex        =   7
            Top             =   495
            Width           =   465
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "RR No."
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
            Left            =   465
            TabIndex        =   3
            Top             =   120
            Width           =   555
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   75
            TabIndex        =   9
            Top             =   870
            Width           =   945
         End
      End
      Begin VB.Label LABALLOWREPRINT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   7230
         TabIndex        =   21
         Top             =   630
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label labStatus 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   6990
         TabIndex        =   20
         Top             =   30
         Width           =   2865
      End
      Begin VB.Label labEDITDetail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10620
         TabIndex        =   36
         Top             =   1125
         Width           =   1155
      End
      Begin VB.Label labid 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   10980
         TabIndex        =   35
         Top             =   990
         Visible         =   0   'False
         Width           =   1140
      End
   End
   Begin VB.PictureBox picVehicleReceving 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6225
      Left            =   0
      ScaleHeight     =   6225
      ScaleWidth      =   13365
      TabIndex        =   37
      Top             =   1950
      Width           =   13365
      Begin VB.CommandButton Command3 
         Caption         =   "Add Specifcation From Model List"
         Height          =   255
         Left            =   6480
         TabIndex        =   120
         ToolTipText     =   "Add Vehicle Specification from Current List"
         Top             =   4290
         Width           =   2685
      End
      Begin VB.PictureBox picVehicleDetails 
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4290
         Left            =   30
         ScaleHeight     =   4290
         ScaleWidth      =   4065
         TabIndex        =   38
         Top             =   30
         Width           =   4065
         Begin VB.ComboBox cboColor 
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
            Left            =   1080
            TabIndex        =   44
            Text            =   "Combo1"
            Top             =   780
            Width           =   2925
         End
         Begin VB.TextBox txtGVW 
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
            Left            =   3090
            MaxLength       =   10
            TabIndex        =   59
            Top             =   2730
            Width           =   915
         End
         Begin VB.TextBox txtPistonDisp 
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
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   62
            Top             =   3120
            Width           =   2925
         End
         Begin VB.TextBox txtFuelUsed 
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
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   58
            Top             =   2730
            Width           =   1185
         End
         Begin VB.TextBox txtEngineNo 
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
            Left            =   1080
            MaxLength       =   25
            TabIndex        =   54
            Top             =   2340
            Width           =   2925
         End
         Begin VB.TextBox txtVINo 
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
            Left            =   1080
            MaxLength       =   25
            TabIndex        =   51
            Top             =   4350
            Width           =   2925
         End
         Begin VB.TextBox txtSerialNo 
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
            Left            =   1080
            MaxLength       =   17
            TabIndex        =   49
            Top             =   1950
            Width           =   2925
         End
         Begin VB.TextBox txtProdNo 
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
            Left            =   1080
            MaxLength       =   25
            TabIndex        =   48
            Top             =   1560
            Width           =   2925
         End
         Begin VB.TextBox txtIgnKey 
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
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   46
            Top             =   1170
            Width           =   2925
         End
         Begin VB.ComboBox cboSource 
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
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   40
            Text            =   "cboSource"
            Top             =   0
            Width           =   2925
         End
         Begin VB.TextBox txtUnit 
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
            Left            =   1080
            TabIndex        =   42
            Top             =   390
            Width           =   2925
         End
         Begin VB.TextBox txtFrameNo 
            BackColor       =   &H000000FF&
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
            Height          =   360
            Left            =   1080
            MaxLength       =   25
            TabIndex        =   56
            Top             =   3510
            Visible         =   0   'False
            Width           =   2925
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   " VIN."
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
            Left            =   600
            TabIndex        =   167
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   " Serial No./"
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
            TabIndex        =   50
            Top             =   1950
            Width           =   900
         End
         Begin VB.Label Label32 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   8
            Left            =   -30
            TabIndex        =   150
            Top             =   2370
            Width           =   150
         End
         Begin VB.Label Label32 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   7
            Left            =   0
            TabIndex        =   149
            Top             =   1920
            Width           =   150
         End
         Begin VB.Label Label32 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   6
            Left            =   570
            TabIndex        =   148
            Top             =   1710
            Width           =   150
         End
         Begin VB.Label Label32 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   5
            Left            =   420
            TabIndex        =   147
            Top             =   1200
            Width           =   150
         End
         Begin VB.Label Label32 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   4
            Left            =   330
            TabIndex        =   146
            Top             =   840
            Width           =   150
         End
         Begin VB.Label Label32 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   3
            Left            =   210
            TabIndex        =   145
            Top             =   30
            Width           =   150
         End
         Begin VB.Label Label15 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "GVW"
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
            Height          =   255
            Left            =   2340
            TabIndex        =   60
            Top             =   2790
            Width           =   1695
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Piston Disp."
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
            Left            =   0
            TabIndex        =   61
            Top             =   3165
            Width           =   1005
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Fuel Used"
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
            Left            =   165
            TabIndex        =   57
            Top             =   2775
            Width           =   825
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Engine No"
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
            Left            =   165
            TabIndex        =   53
            Top             =   2385
            Width           =   840
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "VIN No"
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
            Left            =   450
            TabIndex        =   52
            Top             =   4470
            Width           =   555
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Production No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   465
            Left            =   90
            TabIndex        =   47
            Top             =   1530
            Width           =   930
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "CS #"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   615
            TabIndex        =   45
            Top             =   1260
            Width           =   360
         End
         Begin VB.Label Label4 
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
            Left            =   525
            TabIndex        =   43
            Top             =   870
            Width           =   450
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   390
            TabIndex        =   39
            Top             =   60
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Unit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   660
            TabIndex        =   41
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Frame No"
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
            Left            =   195
            TabIndex        =   55
            Top             =   5160
            Width           =   810
         End
      End
      Begin VB.PictureBox picVehicleRemarks 
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
         Height          =   4260
         Left            =   4170
         ScaleHeight     =   4260
         ScaleWidth      =   4995
         TabIndex        =   96
         Top             =   30
         Width           =   4995
         Begin VB.TextBox txtLTOStatus 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   3030
            TabIndex        =   119
            Top             =   3840
            Width           =   1965
         End
         Begin VB.TextBox txtRemarks1 
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
            Height          =   840
            Left            =   0
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   99
            Top             =   270
            Width           =   2985
         End
         Begin VB.TextBox txtCSR 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   3030
            TabIndex        =   112
            Top             =   3147
            Width           =   1965
         End
         Begin VB.TextBox txtRemarks2 
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
            Height          =   840
            Left            =   0
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   104
            Top             =   1320
            Width           =   2985
         End
         Begin VB.TextBox txtRemarks3 
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
            Height          =   840
            Left            =   0
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   109
            Top             =   2370
            Width           =   2985
         End
         Begin VB.TextBox txtProfile1 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   3030
            TabIndex        =   100
            Top             =   343
            Width           =   1965
         End
         Begin VB.TextBox txtProfile2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   3030
            MaxLength       =   20
            TabIndex        =   110
            Top             =   2446
            Width           =   1965
         End
         Begin VB.TextBox txtProfile4 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   3030
            MaxLength       =   20
            TabIndex        =   106
            Top             =   1745
            Width           =   1965
         End
         Begin VB.TextBox txtProfile3 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   3030
            TabIndex        =   103
            Top             =   1044
            Width           =   1965
         End
         Begin VB.PictureBox picVehicleProfile 
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
            Height          =   990
            Left            =   0
            ScaleHeight     =   990
            ScaleWidth      =   2985
            TabIndex        =   113
            Top             =   3240
            Width           =   2985
            Begin VB.OptionButton optOnShowroom 
               Caption         =   "For Display in Showroom"
               Height          =   210
               Left            =   60
               TabIndex        =   116
               Top             =   495
               Width           =   3165
            End
            Begin VB.OptionButton optWithProsBuyers 
               Caption         =   "Units with Prospective Buyers"
               Height          =   210
               Left            =   60
               TabIndex        =   117
               Top             =   720
               Width           =   3885
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Unknown"
               Height          =   210
               Left            =   60
               TabIndex        =   115
               Top             =   285
               Width           =   3465
            End
            Begin VB.OptionButton optReserved 
               Caption         =   "Unit is Reserved"
               Height          =   210
               Left            =   60
               TabIndex        =   114
               Top             =   60
               Width           =   3465
            End
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LTO Status"
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
            Left            =   3060
            TabIndex        =   118
            Top             =   3550
            Width           =   945
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   0
            TabIndex        =   97
            Top             =   60
            Width           =   810
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "CSR"
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
            Left            =   3120
            TabIndex        =   111
            Top             =   2864
            Width           =   360
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   30
            TabIndex        =   102
            Top             =   1110
            Width           =   810
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   30
            TabIndex        =   107
            Top             =   2160
            Width           =   810
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Gross WT"
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
            Left            =   3060
            TabIndex        =   98
            Top             =   60
            Width           =   840
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "KEY NO"
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
            Left            =   3060
            TabIndex        =   105
            Top             =   1462
            Width           =   630
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "BATTERY"
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
            Left            =   3090
            TabIndex        =   101
            Top             =   761
            Width           =   780
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "TIRES"
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
            Left            =   3060
            TabIndex        =   108
            Top             =   2163
            Width           =   495
         End
      End
      Begin VB.TextBox txtNote 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1650
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   122
         Top             =   4560
         Width           =   9090
      End
      Begin VB.PictureBox picVehiclePricing 
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6540
         Left            =   9090
         ScaleHeight     =   6540
         ScaleWidth      =   4200
         TabIndex        =   63
         Top             =   -30
         Width           =   4200
         Begin VB.TextBox txtFreeSummation 
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
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   1290
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   1530
            Width           =   2715
         End
         Begin MSMask.MaskEdBox txtPurchPrice 
            Height          =   345
            Left            =   1290
            TabIndex        =   65
            Top             =   60
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSubTotalCost 
            Height          =   345
            Left            =   1290
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1170
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   64
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSubisidy 
            Height          =   345
            Left            =   1290
            TabIndex        =   66
            Top             =   435
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtAccSummation 
            Height          =   345
            Left            =   1290
            TabIndex        =   69
            Top             =   795
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFBBody 
            Height          =   345
            Left            =   1275
            TabIndex        =   75
            Top             =   2175
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtAircon 
            Height          =   345
            Left            =   1275
            TabIndex        =   77
            Top             =   2535
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtStereo 
            Height          =   345
            Left            =   1275
            TabIndex        =   79
            Top             =   2910
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCodeAlarm 
            Height          =   345
            Left            =   1275
            TabIndex        =   81
            Top             =   3270
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPullOut 
            Height          =   345
            Left            =   1275
            TabIndex        =   83
            Top             =   3645
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtLto 
            Height          =   345
            Left            =   1275
            TabIndex        =   84
            Top             =   4005
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtTint 
            Height          =   345
            Left            =   1275
            TabIndex        =   87
            Top             =   4380
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSeatCover 
            Height          =   345
            Left            =   1275
            TabIndex        =   89
            Top             =   4740
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMSPlusCard 
            Height          =   345
            Left            =   1275
            TabIndex        =   91
            Top             =   5115
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFloormat 
            Height          =   345
            Left            =   1275
            TabIndex        =   93
            Top             =   5490
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtTotalCost 
            Height          =   375
            Left            =   1275
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   5850
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   64
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total"
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
            Left            =   345
            TabIndex        =   70
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Purch Price"
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
            Left            =   195
            TabIndex        =   64
            Top             =   135
            Width           =   1005
         End
         Begin VB.Label Label3 
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
            Left            =   525
            TabIndex        =   67
            Top             =   555
            Width           =   675
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Free Bies Amount"
            ForeColor       =   &H00000000&
            Height          =   630
            Left            =   135
            TabIndex        =   73
            Top             =   1575
            Width           =   1065
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Accessories"
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
            Left            =   150
            TabIndex        =   68
            Top             =   885
            Width           =   1080
         End
         Begin VB.Line Line1 
            BorderStyle     =   3  'Dot
            X1              =   75
            X2              =   4050
            Y1              =   1995
            Y2              =   1995
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Cost"
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
            Left            =   300
            TabIndex        =   94
            Top             =   5970
            Width           =   855
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "FB Body"
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
            Left            =   525
            TabIndex        =   74
            Top             =   2250
            Width           =   675
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Aircon"
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
            Left            =   645
            TabIndex        =   76
            Top             =   2595
            Width           =   555
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Stereo"
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
            Left            =   630
            TabIndex        =   78
            Top             =   2940
            Width           =   570
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Code Alarm"
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
            Left            =   210
            TabIndex        =   80
            Top             =   3330
            Width           =   990
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Pull Out"
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
            Left            =   540
            TabIndex        =   82
            Top             =   3720
            Width           =   660
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LTO"
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
            Left            =   855
            TabIndex        =   85
            Top             =   4125
            Width           =   345
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Tint"
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
            Left            =   885
            TabIndex        =   86
            Top             =   4470
            Width           =   315
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Seat Cover"
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
            Left            =   270
            TabIndex        =   88
            Top             =   4770
            Width           =   930
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Other 1"
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
            Index           =   0
            Left            =   570
            TabIndex        =   90
            Top             =   5175
            Width           =   630
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Floormat"
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
            Left            =   450
            TabIndex        =   92
            Top             =   5520
            Width           =   750
         End
      End
      Begin VB.Label Label40 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Note and Specifications"
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
         Height          =   315
         Left            =   120
         TabIndex        =   121
         Top             =   4320
         Width           =   3315
      End
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   -60
      ScaleHeight     =   345
      ScaleWidth      =   4350
      TabIndex        =   158
      Top             =   8190
      Width           =   4350
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " APJ #:"
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
         Index           =   2
         Left            =   60
         TabIndex        =   161
         Top             =   0
         Width           =   825
      End
      Begin VB.Label labAPJ 
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
         Left            =   900
         TabIndex        =   160
         Top             =   0
         Width           =   1065
      End
      Begin VB.Label labDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   2040
         TabIndex        =   159
         Top             =   60
         Width           =   2355
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   1980
         Top             =   0
         Width           =   2355
      End
   End
   Begin VB.PictureBox picBottoms 
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   30
      ScaleHeight     =   930
      ScaleWidth      =   18825
      TabIndex        =   123
      Top             =   8520
      Width           =   18825
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   5490
         Top             =   240
      End
      Begin Crystal.CrystalReport rptMRR 
         Left            =   5490
         Top             =   315
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Vehicle Receiving Report"
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
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
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
         Left            =   11415
         ScaleHeight     =   885
         ScaleWidth      =   1800
         TabIndex        =   139
         Top             =   0
         Width           =   1800
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            CausesValidation=   0   'False
            Height          =   795
            Left            =   945
            MouseIcon       =   "MRRINV.frx":18FA
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":1A4C
            Style           =   1  'Graphical
            TabIndex        =   140
            ToolTipText     =   "Cancel"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   255
            MouseIcon       =   "MRRINV.frx":1D8A
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":1EDC
            Style           =   1  'Graphical
            TabIndex        =   141
            ToolTipText     =   "Save this Record"
            Top             =   0
            Width           =   705
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
         Height          =   945
         Left            =   2280
         ScaleHeight     =   945
         ScaleWidth      =   11145
         TabIndex        =   124
         Top             =   0
         Width           =   11145
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   10050
            MouseIcon       =   "MRRINV.frx":222C
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":237E
            Style           =   1  'Graphical
            TabIndex        =   138
            ToolTipText     =   "Exit Window"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   795
            Left            =   9360
            MouseIcon       =   "MRRINV.frx":26E4
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":2836
            Style           =   1  'Graphical
            TabIndex        =   137
            ToolTipText     =   "Print this Record"
            Top             =   0
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
            Left            =   8670
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "MRRINV.frx":2B9C
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":2CEE
            Style           =   1  'Graphical
            TabIndex        =   136
            ToolTipText     =   "Cancel this Transaction"
            Top             =   0
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
            Left            =   7980
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "MRRINV.frx":3028
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":317A
            Style           =   1  'Graphical
            TabIndex        =   135
            ToolTipText     =   "Post this Transaction"
            Top             =   0
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
            Left            =   7290
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "MRRINV.frx":349F
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":35F1
            Style           =   1  'Graphical
            TabIndex        =   134
            ToolTipText     =   "Unpost this Transaction"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   6600
            MouseIcon       =   "MRRINV.frx":3936
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":3A88
            Style           =   1  'Graphical
            TabIndex        =   131
            ToolTipText     =   "Edit Selected Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   5910
            MouseIcon       =   "MRRINV.frx":3DE4
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":3F36
            Style           =   1  'Graphical
            TabIndex        =   133
            ToolTipText     =   "Add Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdLast 
            Caption         =   "Last"
            Height          =   795
            Left            =   5220
            MouseIcon       =   "MRRINV.frx":4249
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":439B
            Style           =   1  'Graphical
            TabIndex        =   132
            ToolTipText     =   "Move to Last Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdFirst 
            Caption         =   "First"
            Height          =   795
            Left            =   4530
            MouseIcon       =   "MRRINV.frx":46EB
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":483D
            Style           =   1  'Graphical
            TabIndex        =   130
            ToolTipText     =   "Move to First Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   3840
            MouseIcon       =   "MRRINV.frx":4B9B
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":4CED
            Style           =   1  'Graphical
            TabIndex        =   129
            ToolTipText     =   "Find a Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   3150
            MouseIcon       =   "MRRINV.frx":4FE7
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":5139
            Style           =   1  'Graphical
            TabIndex        =   128
            ToolTipText     =   "Move to Next Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   2460
            MouseIcon       =   "MRRINV.frx":5491
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":55E3
            Style           =   1  'Graphical
            TabIndex        =   126
            ToolTipText     =   "Move to Previous Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3420
            MouseIcon       =   "MRRINV.frx":5942
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":5A94
            Style           =   1  'Graphical
            TabIndex        =   127
            ToolTipText     =   "Delete Selected Record"
            Top             =   0
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   795
            Left            =   1770
            MouseIcon       =   "MRRINV.frx":5DBF
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":5F11
            Style           =   1  'Graphical
            TabIndex        =   125
            ToolTipText     =   "Move to Previous Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdServiceVehicle 
            Caption         =   "&Service Vehicle Info"
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
            Left            =   660
            MouseIcon       =   "MRRINV.frx":648C
            MousePointer    =   99  'Custom
            Picture         =   "MRRINV.frx":65DE
            Style           =   1  'Graphical
            TabIndex        =   163
            ToolTipText     =   "Save Changes"
            Top             =   0
            Width           =   1125
         End
      End
   End
   Begin VB.Label LABVEHREGISTER 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   8640
      TabIndex        =   166
      Top             =   8190
      Width           =   4485
   End
   Begin VB.Label labInventoryStatus 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   4290
      TabIndex        =   162
      Top             =   8190
      Width           =   4335
   End
End
Attribute VB_Name = "frmSMIS_Trans_MRR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMRRINV                                                          As ADODB.Recordset
Dim AddorEdit                                                         As String
Attribute AddorEdit.VB_VarUserMemId = 1073938435
Dim Tutal                                                             As Currency
Attribute Tutal.VB_VarUserMemId = 1073938436
Dim WithEvents SearchMaster                                           As frmSMIS_Mis_SearchMaster
Attribute SearchMaster.VB_VarHelpID = -1

Function DetectATMT(strx)
    Dim i                                                             As Integer
    Dim ax
    ax = Split(strx)

    For i = 1 To UBound(ax)
        If InStr(1, ax(i), "MT") > 0 Then
            DetectATMT = "MT"
            Exit Function
        ElseIf InStr(1, ax(i), "M/T") > 0 Then
            DetectATMT = "MT"
            Exit Function
        End If
    Next


    For i = 1 To UBound(ax)
        If InStr(1, ax(i), "AT") > 0 Then
            DetectATMT = "AT"
            Exit Function
        ElseIf InStr(1, ax(i), "A/T") > 0 Then
            DetectATMT = "AT"
            Exit Function
        End If
    Next

    DetectATMT = ""

    Erase ax
End Function

Function GetClassCode() As String
    Dim temprs                                                        As ADODB.Recordset

    If cboClass.ListIndex <> -1 Then

        Set temprs = gconDMIS.Execute("SELECT CODE FROM SMIS_VehiclesClass Where ID= " & cboClass.ItemData(cboClass.ListIndex))

        If Not (temprs.EOF Or temprs.BOF) Then
            GetClassCode = Null2String(temprs!CODE)
        End If

        Set temprs = Nothing

    Else
        GetClassCode = vbNullString
    End If
End Function

Function ReturnAccountsPayableInformation(vRR As String, VCOND As String) As String()
    Dim rsKUTO                                                        As New ADODB.Recordset
    Dim VTMP(1)                                                       As String
    Set rsKUTO = gconDMIS.Execute("SELECT VOUCHERNO,STATUS FROM AMIS_JOURNAL_HD WHERE INVOICENO = '" & vRR & "' AND JTYPE = 'APJ'")
    If Not (rsKUTO.BOF And rsKUTO.EOF) Then
        If Null2String(rsKUTO!STATUS) = "P" Then
            VTMP(0) = Null2String(rsKUTO!VOUCHERNO)
            VTMP(1) = "POSTED IN AMIS"
        Else
            VTMP(0) = Null2String(rsKUTO!VOUCHERNO)
            VTMP(1) = "IMPORTED IN AMIS"
        End If
    Else
        Set rsKUTO = New ADODB.Recordset
        Set rsKUTO = gconDMIS.Execute("SELECT VOUCHERNO,STATUS FROM AMIS_JOURNAL_HD WHERE INVOICENO = '" & VCOND & "' AND JTYPE = 'APJ'")
        If Not (rsKUTO.BOF And rsKUTO.EOF) Then
            If Null2String(rsKUTO!STATUS) = "P" Then
                VTMP(0) = Null2String(rsKUTO!VOUCHERNO)
                VTMP(1) = "POSTED IN AMIS"
            Else
                VTMP(0) = Null2String(rsKUTO!VOUCHERNO)
                VTMP(1) = "IMPORTED IN AMIS"
            End If
        Else
            VTMP(0) = ""
            VTMP(1) = ""
            ReturnAccountsPayableInformation = VTMP
        End If
    End If

    ReturnAccountsPayableInformation = VTMP
    Set rsKUTO = Nothing
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

Sub InitData()



    FillCombo "SELECT DISTINCT DESCRIPT from ALL_MODEL ", -1, 0, cboModelDescript


    FillCombo "SELECT COLOR_DESC from ALL_COLOR", -1, 0, cboColor
    FillCombo "SELECT ID, CLASSNAME from SMIS_VehiclesClass", 0, 1, cboClass

    Combo_Loadval cboSource, gconDMIS.Execute("SELECT rtrim(ltrim(DEALERCODE))DealerCode  FROM CSMS_SELLINGDEALER order by dealercode asc")
    SetComboWidth cboClass, 250
    SetComboMaxLength cboTransmission, 2

End Sub

Sub initMemvars()

    txtVeh_register_CustCode = ""
    txtVeh_register_CustName = ""
    LABVEHREGISTER = ""
    labStatus = ""
    labDetails = ""
    labAPJ = ""
    LABALLOWREPRINT = ""
    labInventoryStatus = ""
    txtCode.Text = ""
    txtPO = ""
    cboModelDescript.Text = ""
    TXTMAKE.Text = "Mitsubishi"
    TXTMAKE.Locked = True
    txtModel.Text = ""
    txtYeer.Text = ""
    txtref_PONO = ""
    txtDRNO = ""
    cboSource = ""

    txtUnit.Text = ""
    txtIgnKey.Text = ""
    txtProdNo.Text = ""
    txtSerialNo.Text = ""
    txtVINO.Text = ""
    txtEngineNo.Text = ""
    txtFuelUsed.Text = ""
    txtPistonDisp.Text = ""
    txtGVW.Text = ""
    txtPurchPrice.Text = "0.00"
    txtSubisidy.Text = "0.00"
    txtFBBody.Text = "0.00"
    txtAircon.Text = "0.00"
    txtStereo.Text = "0.00"
    txtCodeAlarm.Text = "0.00"
    txtPullOut.Text = "0.00"
    txtLto.Text = "0.00"
    txtTint.Text = "0.00"
    txtSeatCover.Text = "0.00"
    txtMSPlusCard.Text = "0.00"
    txtFloormat.Text = "0.00"
    txtDateReceived.Text = ""
    txtDateReleased.Text = ""
    txtProfile1.Text = ""
    txtProfile2.Text = ""
    txtProfile3.Text = ""
    txtProfile4.Text = ""
    txtPullOutDate.Text = ""
    txtRemarks1.Text = ""
    txtRemarks2.Text = ""
    txtRemarks3.Text = ""
    txtLTOStatus.Text = ""
    txtCSR.Text = ""
    txtNote.Text = ""
    txtModelCode = ""

    txtAccSummation = "0.00"
    cboColor = ""
    txtFrameNo = ""
    optOnShowroom.Value = True
    optWithProsBuyers.Value = False
End Sub

Sub loadthespec()
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset

    SQL = "SELECT spec From ALL_MODEL where descript='" & cboModelDescript.Text & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.BOF And Not RS.EOF Then

        txtNote.Text = Null2String(RS!spec)

    End If
    Set RS = Nothing
End Sub

Sub SearchID(XXX)

    Dim varBookMark                                                   As Variant
    varBookMark = rsMRRINV.Bookmark
    rsMRRINV.MoveFirst
    rsMRRINV.Find "id = " & XXX

    If (rsMRRINV.BOF = True) Or (rsMRRINV.EOF = True) Then
        MsgBox "Record not found"
        rsMRRINV.Bookmark = varBookMark
    End If

    StoreMemVars
End Sub

Sub SetClass()
    If rsMRRINV.EOF Or rsMRRINV.BOF Then
        Exit Sub
    End If
    Dim temprs                                                        As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT ClassName FROM SMIS_VehiclesClass Where Code= " & N2Str2Null(rsMRRINV!Class))

    If Not (temprs.EOF Or temprs.BOF) Then
        cboClass.ListIndex = SelectCombo(cboClass, Null2String(temprs!ClassName))
    End If

    Set temprs = Nothing


End Sub

Private Sub cmdRefresh_Click()
    If Not (rsMRRINV.EOF Or rsMRRINV.BOF) Then
        rsRefresh
        rsMRRINV.Find ("ID=" & labid)
        StoreMemVars
    End If
End Sub

Private Sub cboModelDescript_Change()
    If AddorEdit = "" Then: Exit Sub
    If RTrim(LTrim(cboModelDescript)) = "" Then: Exit Sub
    Dim temprs                                                        As ADODB.Recordset
    Dim rsModelCode                                                   As ADODB.Recordset
    Dim prodno
    Set temprs = gconDMIS.Execute("Select MODEL,spec,COSTPRICE from ALL_MODEL where descript=" & N2Str2Null(cboModelDescript))

    If Not (temprs.BOF Or temprs.EOF) Then
        txtModel = Null2String(temprs!Model)
        txtNote = Null2String(temprs!spec)
        If AddorEdit = "ADD" Then
            txtPurchPrice = Null2String(temprs!costprice)
        End If
        Set rsModelCode = gconDMIS.Execute("select CODE FROM ALL_ModelCode where description=" & N2Str2Null(txtModel))
        If Not rsModelCode.EOF Or Not rsModelCode.BOF Then
            txtModelCode.Text = Null2String(rsModelCode!CODE)
        End If
    End If
    txtUnit.Text = cboModelDescript.Text
    If (AddorEdit = "ADD") Or (AddorEdit = "EDIT" And cboTransmission = "") Then
        cboTransmission.Text = DetectATMT(cboModelDescript)
    End If


    txtProdNo = ""

    If AddorEdit = "ADD" Then
        Dim RSPRODNO                                                  As ADODB.Recordset
        Set RSPRODNO = gconDMIS.Execute("SELECT top 1  * FROM SMIS_MrrInv_Table  WHERE isnull(profile1,'')<>'' and  descript=" & N2Str2Null(cboModelDescript))
        If Not (RSPRODNO.EOF Or RSPRODNO.BOF) Then
            prodno = Split(Null2String(RSPRODNO!prodno), "-")
            If UBound(prodno) >= 1 Then
                txtProdNo = prodno(0)
            End If
            If LTrim(RTrim(COMPANY_CODE)) = "HBK" Then
                txtProfile1 = Null2String(RSPRODNO!profile1)
                txtProfile2 = Null2String(RSPRODNO!profile2)
                txtProfile3 = Null2String(RSPRODNO!profile3)
                txtProfile4 = Null2String(RSPRODNO!profile4)
            End If
        End If
        Set RSPRODNO = Nothing
    End If
    Set temprs = Nothing
    Set rsModelCode = Nothing
End Sub

Private Sub cboModelDescript_Click()
    cboModelDescript_Change
    loadthespec
End Sub

Private Sub cboModelDescript_GotFocus()
    VBComBoBoxDroppedDown cboModelDescript
End Sub

Private Sub cboSource_LostFocus()
    '    cboSource.ListIndex = SelectCombo(cboSource, cboSource)
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "VEHICLE RECIEVING") = False Then Exit Sub
    On Error GoTo ErrorCode:
    AddorEdit = "ADD"
    initMemvars
    txtCode = GenerateCode("SMIS_MRRINV_TABLE", "CODE", "000000")
    txtPullOutDate.Enabled = False: txtDateReceived.Enabled = False

    picVehicleReceving.Enabled = True
    picTops.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    picModelDetails.Enabled = True

    picRefHeader.Enabled = True
    picVehiclePricing.Enabled = True
    picVehicleDetails.Enabled = True
    txtDateReceived = LOGDATE
    txtPullOutDate = LOGDATE
    'txtCode = GenerateCode("SMIS_MRRINV", "CODE", "000000")
    On Error Resume Next
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdAddFromPO_Click()
    'UDPATING CODE      :   AXP-06132007149
    SearchMaster.SearchForPO
    SearchMaster.Show 1
End Sub

Private Sub cmdCancel_Click()
    cboColor.Enabled = True
    txtIgnKey.Enabled = True
    txtProdNo.Enabled = True
    txtSerialNo.Enabled = True
    cmdAddFromPO.Enabled = True
    AddorEdit = ""
    picTops.Enabled = False: picAdds.Visible = True: picSaves.Visible = False: picVehicleReceving.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "VEHICLE RECIEVING") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox(" Are You Sure You Want To Cancel This Transaction", vbYesNo + vbQuestion) = vbNo Then: Exit Sub
    If Null2String(rsMRRINV!ISTATUS) <> "O" Then
        MsgBox " Record Cannot Be Cancelled Vehicle Information is in Use", vbExclamation
        Exit Sub
    End If
    SQL_STATEMENT = "update SMIS_MrrInv_Table set  status='C' where id = " & labid.Caption
    '**********NEW LOG AUDIT************
    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "C", "VEHICLE RECIEVING", SQL_STATEMENT, Null2String(labid), "", "RR No:" & txtCode, "", ""
    '**********************************
    gconDMIS.Execute ("update SMIS_PO set  DateReceived=NULL where PO_NO= " & Null2String(rsMRRINV!PONO))

    MessagePop RecSaveOk, "Transaction Cancelled", "Record Sucessfully Cancelled", 1000, 2
    'LogAudit "C", "VEHICLE RECEVING", cboModelDescript & " CS: " & txtIgnKey & " DATE RECEVIED:" & txtDateReceived
    rsRefresh
    rsMRRINV.Find ("ID=" & labid)
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "VEHICLE RECIEVING") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If ShowConfirmDelete = True Then
        SQL_STATEMENT = "delete from SMIS_MrrInv where id = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT---------------------------------------------------
            Call NEW_LogAudit("X", "VEHICLE RECIEVING", SQL_STATEMENT, labid, "", "RR NO: " & txtCode, "", "")
        'NEW LOG AUDIT---------------------------------------------------
        
        ShowDeletedMsg
        rsRefresh
        StoreMemVars
    End If
    
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "VEHICLE RECIEVING") = False Then Exit Sub
    txtPullOutDate.Enabled = False
    txtDateReceived.Enabled = False
    On Error GoTo ErrorCode:
    AddorEdit = "EDIT"
    picTops.Enabled = True
    picVehicleReceving.Enabled = True
    picSaves.Visible = True
    'UPDATED BY: JUN------------------------------------------------------------------------------------------------
    'DATE UPDATE: 10142008
    'DESCRIPTION: DISABLE THE BUTTON IF THE MODE IS EDITING RECEIVEING SO THAT OTHER PENDING PO WILL NOT BE AFFECTED
    cmdAddFromPO.Enabled = False
    'UPDATED BY: JUN------------------------------------------------------------------------------------------------
    picAdds.Visible = False
    If labInventoryStatus <> "** AVAILABLE/OPEN**" Then
        MessagePop RecLocekd, "Vehicle is In use", "Editing Is Limited. Vehicle is Already In use"
    End If
    On Error Resume Next
    cboSource.SetFocus

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    frmSMIS_SearchVehicleInfo.Show
End Sub

Private Sub cmdFirst_Click()
    On Error GoTo ErrorCode:
    If rsMRRINV.BOF Then
        ShowFirstRecordMsg
    Else
        rsMRRINV.MoveFirst
        StoreMemVars
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdLast_Click()
    On Error GoTo ErrorCode:
    If rsMRRINV.EOF Then
        ShowFirstRecordMsg
    Else
        rsMRRINV.MoveLast
        StoreMemVars
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:
    rsMRRINV.MoveNext
    If rsMRRINV.EOF Then
        rsMRRINV.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "VEHICLE RECIEVING") = False Then Exit Sub

    Dim rsCusVeh1                                                     As ADODB.Recordset
    Set rsCusVeh1 = gconDMIS.Execute("Select CUSCDE,NIYM,INVOICENO From CSMS_CUSVEH WHERE UPPER(MAKE) = 'Mitsubishi' AND VCOND_NO='" & txtIgnKey & "'")
    If (rsCusVeh1.EOF Or rsCusVeh1.BOF) Then
        MsgBox "Vehicle Has not Been Register to Service Department! " & vbCrLf & "Please Register Vehicle Prior to Posting", vbCritical, "Registration Required"
        Exit Sub

    End If

    On Error GoTo ErrorCode:
    Dim SQL                                                           As String
    If MsgBox("Are You Sure You Want To Post This Transaction", vbInformation + vbYesNo) = vbNo Then: Exit Sub
    'gconDMIS.Execute "update SMIS_MrrInv_Table set status='P' where id = " & labid.Caption

    SQL_STATEMENT = "update SMIS_MrrInv_Table set status='P' where id = " & labid.Caption
    gconDMIS.Execute (SQL_STATEMENT)
    '**********NEW LOG AUDIT**************
        NEW_LogAudit "P", "VEHICLE RECIEVING", SQL_STATEMENT, labid, "", "RR No:" & txtCode, "", ""
    '*************************************

    MessagePop RecSaveOk, "Transaction Posted", "Record Sucessfully Posted", 1000, 2
    'LogAudit "P", "VEHICLE RECEVING", cboModelDescript & " CS: " & txtIgnKey & " DATE RECEVIED:" & txtDateReceived
    rsRefresh
    rsMRRINV.Find ("ID=" & labid)
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:
    rsMRRINV.MovePrevious
    If rsMRRINV.BOF Then
        rsMRRINV.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "VEHICLE RECIEVING") = False Then Exit Sub
    If LABALLOWREPRINT <> "" Then
        If AllowReprint("VEHICLE RECIEVING") = False Then Exit Sub
    End If
    On Error GoTo ErrorCode:
    Screen.MousePointer = 11
    LoadSignatories ("RECIEVING REPORT")
    If COMPANY_CODE = "HBK" Then
        rptMRR.Formulas(0) = "Preparedby= '" & Null2String(PreparedBy) & "'"
        rptMRR.Formulas(1) = "Recievedby= '" & Null2String(ApprovedBy) & "'"
        rptMRR.Formulas(2) = "Deliveredby= '" & Null2String(SalesDispatcher) & "'"
        rptMRR.Formulas(3) = "CheckedBy='" & Null2String(CheckedBy) & "'"
        
        'rptMRR.Formulas(4) = "Recievedby= '" & Null2String(ApprovedBy) & "'"
        'rptMRR.Formulas(5) = "Deliveredby= '" & Null2String(SalesDispatcher) & "'"
        'rptMRR.Formulas(6) = "CheckedBy='" & Null2String(CheckedBy) & "'"
    Else
        rptMRR.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptMRR.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptMRR.Formulas(2) = "PreparedBy = '" & PreparedBy & "'"
    End If
    PrintSQLReport rptMRR, SMIS_REPORT_PATH & "MRR.rpt", "{MRRINV.ID} = " & labid, DMIS_REPORT_Connection, 1
    SQL_STATEMENT = "UPDATE SMIS_MRRINV_TABLE  SET PRINTED=1 WHERE ID=" & labid
    gconDMIS.Execute SQL_STATEMENT
    
    '*****NEW LOG AUDIT************
    NEW_LogAudit "V", "VEHICLE RECIEVING", SQL_STATEMENT, labid, "", "RR No:" & txtCode, "", ""
    '*****************************

    rsRefresh
    rsMRRINV.Find ("ID=" & labid)
    StoreMemVars
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    Dim lngcount                                                      As Integer
    If NumericVal(txtTotalCost) <= 0 Then
        MsgBox "Warning!Total Vehicle Cost Cannot Be Zero", vbInformation
        Exit Sub
    End If

    If Trim(cboSource) = "" Then
        MsgBox "Warning!Source of Vehicle Cannot Be Blank", vbInformation
        cboSource.SetFocus
        Exit Sub
    End If

    If Trim(TXTMAKE) = "" Then
        ShowIsRequiredMsg "Make"
        On Error Resume Next
        TXTMAKE.SetFocus
        Exit Sub
    End If

    If IsDate(txtPullOutDate) = False Then
        ShowIsRequiredMsg "Invalid PullOut Date"
        On Error Resume Next
        txtPullOutDate.SetFocus
        Exit Sub
    End If
    If RTrim(LTrim(txtCode.Text)) = "" Then
        ShowIsRequiredMsg "MRR Code"
        On Error Resume Next
        txtCode.SetFocus
        Exit Sub
    End If
    If RTrim(LTrim(txtEngineNo.Text)) = "" Then
        ShowIsRequiredMsg "Engine No"
        On Error Resume Next
        txtEngineNo.SetFocus
        Exit Sub
    End If
    If RTrim(LTrim(txtIgnKey.Text)) = "" Then
        ShowIsRequiredMsg "Ignition Key"
        On Error Resume Next
        txtIgnKey.SetFocus
        Exit Sub
    End If
    If RTrim(LTrim(txtSerialNo.Text)) = "" Then
        ShowIsRequiredMsg "Serial. No."
        On Error Resume Next
        txtSerialNo.SetFocus
        Exit Sub
    End If
    If RTrim(LTrim(txtProdNo.Text)) = "" Then
        ShowIsRequiredMsg "Prod. No."
        On Error Resume Next
        txtProdNo.SetFocus
        Exit Sub
    End If

    If RTrim(LTrim(txtVINO.Text)) = "" Then
        ShowIsRequiredMsg "VIN Number."
        On Error Resume Next
        txtVINO.SetFocus
        Exit Sub
    End If
    If IsDate(txtDateReceived.Text) = False Or txtDateReceived.Text = "" Then
        MsgBoxXP "Invalid Date Received... Pls. input the Invoice Date Properly!", "Error", XP_OKOnly, msg_Critical
        On Error Resume Next
        txtDateReceived.SetFocus
        Exit Sub
    End If
    If txtModel.Text = "" Then
        ShowIsRequiredMsg "Model"
        On Error Resume Next
        txtModel.SetFocus
        Exit Sub
    End If
    If cboModelDescript.Text = "" Then
        ShowIsRequiredMsg "Description"
        On Error Resume Next
        cboModelDescript.SetFocus
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lngcount = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_MrrInv_table WHERE upper(LTrim(RTrim(code)))=" & N2Str2Null(UCase(LTrim(RTrim(txtCode))))).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lngcount >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "MRR Number Already Exist"
            txtCode.SetFocus
            Exit Sub
        End If
    Else
        If lngcount >= 1 And Null2String(UCase(LTrim(RTrim(rsMRRINV!CODE)))) <> UCase(LTrim(RTrim(txtCode))) Then
            MessagePop RecSaveWarning, "Duplicate Record", "MRR Number Already Exist"
            txtCode.SetFocus
            Exit Sub
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lngcount = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_MrrInv WHERE VINO=" & N2Str2Null(txtVINO)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lngcount >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "VIN Number Already Exist"
            txtVINO.SetFocus
            Exit Sub
        End If
    Else
        If lngcount >= 1 And UCase(Null2String(rsMRRINV!VINO)) <> UCase(txtVINO) Then
            MessagePop RecSaveWarning, "Duplicate Record", "VIN Number Already Exist"
            txtVINO.SetFocus
            Exit Sub
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lngcount = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_MrrInv WHERE IGNKEY=" & N2Str2Null(txtIgnKey)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lngcount >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Conduction Sticker  Already Exist"
            txtIgnKey.SetFocus
            Exit Sub
        End If
    Else
        If lngcount >= 1 And UCase(Null2String(rsMRRINV!ignkey)) <> UCase(txtIgnKey) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Conduction Sticker  Already Exist.. Please Use Another One"
            txtIgnKey.SetFocus
            Exit Sub
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lngcount = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_MrrInv WHERE Prodno=" & N2Str2Null(txtProdNo)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lngcount >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Product Number Of Such Code Already Exist"
            txtProdNo.SetFocus
            Exit Sub
        End If
    Else
        If lngcount >= 1 And UCase(Null2String(rsMRRINV!prodno)) <> UCase(txtProdNo) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Product Number Of Such Code Already Exist. Please Use Another one"
            txtProdNo.SetFocus
            Exit Sub
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'UPDATED BY: JUN
    'DATE UPDATED: 11042008
    'DESCRIPTION: CONTROL DUPLICATE REFERENCE INVOICE NO
    lngcount = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_MrrInv WHERE REFPONO=" & N2Str2Null(txtref_PONO)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lngcount >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Reference Invoice Already Exist"
            txtref_PONO.SetFocus
            Exit Sub
        End If
    Else
        If lngcount >= 1 And UCase(Null2String(rsMRRINV!refPONO)) <> UCase(txtref_PONO) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Reference Invoice Already Exist"
            txtref_PONO.SetFocus
            Exit Sub
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim vtxtCode, vtxtdescript, vtxtmake, vtxtmodel, vtxtTransmission, vtxtModelCode, vcboclass, vtxtyeer, vcbosource, vtxtunit, vcbocolor As String
    Dim vtxtignkey, vtxtprodno, vtxtserialno, vtxtVINo, vtxtengineno, vtxtfuelused, vtxtpistondisp, vtxtgvw, vtxtpurchprice, vtxtframeno As String
    Dim vtxtmmpcsubs, vtxtfbbody, vtxtaircon, vtxtstereo, vtxtcodealarm, vtxtpullout, vtxtlto, vtxttint, vtxtseatcover, vtxtmspluscard, vtxtfloormat As Double
    Dim vtxtdatereceived, vtxtdatereleased                            As String
    Dim vtxtprofile1, vtxtprofile2, vtxtprofile3, vtxtprofile4        As String
    Dim vtxtpulloutdate, voptonshowroom, voptwithprosbuyers, vtxtremarks1, vtxtremarks2, vtxtremarks3 As String
    Dim vtxtltostatus, vtxtcsr, vtxtnote, vtxtdrno, vtxtrefpono, vtxtpono As String

    vtxtCode = N2Str2Null(txtCode)
    vtxtdescript = N2Str2Null(cboModelDescript)
    vtxtmake = N2Str2Null(TXTMAKE)
    vtxtmodel = N2Str2Null(txtModel)
    vtxtModelCode = N2Str2Null(txtModelCode)
    vtxtTransmission = N2Str2Null(cboTransmission)
    vcboclass = N2Str2Null(GetClassCode)
    vtxtyeer = N2Str2Null(txtYeer)
    vcbosource = N2Str2Null(cboSource)
    vtxtunit = N2Str2Null(txtUnit)
    vcbocolor = N2Str2Null(cboColor)

    vtxtignkey = N2Str2Null(txtIgnKey)
    vtxtprodno = N2Str2Null(txtProdNo)
    vtxtserialno = N2Str2Null(txtSerialNo)
    vtxtVINo = N2Str2Null(txtVINO)
    vtxtengineno = N2Str2Null(txtEngineNo)
    vtxtfuelused = N2Str2Null(txtFuelUsed)
    vtxtpistondisp = N2Str2Null(txtPistonDisp)
    vtxtgvw = N2Str2Null(txtGVW)
    vtxtpurchprice = N2Str2Zero(txtPurchPrice)
    vtxtframeno = N2Str2Null(txtFrameNo)

    vtxtmmpcsubs = N2Str2Zero(txtSubisidy)
    vtxtfbbody = N2Str2Zero(txtFBBody)
    vtxtaircon = N2Str2Zero(txtAircon)
    vtxtstereo = N2Str2Zero(txtStereo)
    vtxtcodealarm = N2Str2Zero(txtCodeAlarm)
    vtxtpullout = N2Str2Zero(txtPullOut)
    vtxtlto = N2Str2Zero(txtLto)
    vtxttint = N2Str2Zero(txtTint)
    vtxtseatcover = N2Str2Zero(txtSeatCover)
    vtxtmspluscard = N2Str2Zero(txtMSPlusCard)
    vtxtfloormat = N2Str2Zero(txtFloormat)
    vtxtdatereceived = N2Date2Null(txtDateReceived)
    vtxtdatereleased = N2Date2Null(txtDateReleased)
    vtxtpulloutdate = N2Date2Null(txtPullOutDate)
    vtxtprofile1 = N2Str2Null(txtProfile1)
    vtxtprofile2 = N2Str2Null(txtProfile2)
    vtxtprofile3 = N2Str2Null(txtProfile3)
    vtxtprofile4 = N2Str2Null(txtProfile4)
    vtxtremarks1 = N2Str2Null(txtRemarks1)
    vtxtremarks2 = N2Str2Null(txtRemarks2)
    vtxtremarks3 = N2Str2Null(txtRemarks3)
    vtxtltostatus = N2Str2Null(txtLTOStatus)
    vtxtcsr = N2Str2Null(txtCSR)
    vtxtnote = N2Str2Null(txtNote)
    vtxtdrno = N2Str2Null(txtDRNO)
    vtxtpono = N2Str2Null(txtPO)
    vtxtrefpono = N2Str2Null(txtref_PONO)


    If optOnShowroom.Value = True = True Then
        voptonshowroom = "'Y'"
    Else
        voptonshowroom = "'N'"
    End If
    If optWithProsBuyers.Value = True Then
        voptwithprosbuyers = "'Y'"
    Else
        voptwithprosbuyers = "'N'"
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "INSERT INTO SMIS_MRRINV_TABLE" & _
                      " (PONO, TRANSMISSION, FRAMENO,ISTATUS, CODE,DESCRIPT,MAKE,MODEL,MODELCODE, CLASS,YEER,SOURCE,UNIT,COLOR,IGNKEY,PRODNO,SERIALNO," & _
                        "VINO,ENGINENO,FUELUSED,PISTONDISP,GVW,PURCHPRICE,MMPCSUBS," & _
                        "FBBODY,AIRCON,STEREO,CODEALARM,PULLOUT,LTO,TINT,SEATCOVER,MSPLUS,FLOORMAT,DATERECEIVED,DATERELEASED,PROFILE1,PROFILE2,PROFILE3,PROFILE4,PULLOUTDATE,REMARKS1,REMARKS2,REMARKS3,LTOSTATUS,CSR,NOTES,ONSHOWROOM,WITHPROSBUYERS, REFPONO,DRNO)" & _
                      " VALUES (" & vtxtpono & "," & vtxtTransmission & "," & vtxtframeno & ", 'O'," & vtxtCode & ", " & vtxtdescript & ", " & vtxtmake & ", " & vtxtmodel & ", " & vtxtModelCode & ", " & vcboclass & "," & _
                      " " & vtxtyeer & ", " & vcbosource & ", " & vtxtunit & "," & _
                      " " & vcbocolor & ", " & vtxtignkey & ", " & vtxtprodno & ", " & vtxtserialno & "," & _
                      " " & vtxtVINo & ", " & vtxtengineno & ", " & vtxtfuelused & "," & _
                      " " & vtxtpistondisp & ", " & vtxtgvw & ", " & vtxtpurchprice & ", " & vtxtmmpcsubs & ", " & vtxtfbbody & "," & _
                      " " & vtxtaircon & ", " & vtxtstereo & ", " & vtxtcodealarm & ", " & vtxtpullout & "," & _
                      " " & vtxtlto & ", " & vtxttint & ", " & vtxtseatcover & ", " & vtxtmspluscard & ", " & vtxtfloormat & ", " & vtxtdatereceived & _
                        ", " & vtxtdatereleased & ", " & vtxtprofile1 & ", " & vtxtprofile2 & ", " & vtxtprofile3 & ", " & vtxtprofile4 & ", " & vtxtpulloutdate & ", " & vtxtremarks1 & "," & vtxtremarks2 & "," & vtxtremarks3 & "," & vtxtltostatus & "," & vtxtcsr & "," & vtxtnote & "," & voptonshowroom & "," & voptwithprosbuyers & "," & vtxtrefpono & "," & vtxtdrno & ")"
        gconDMIS.Execute (SQL_STATEMENT)

        '****************NEW LOG AUDIT***********
        NEW_LogAudit "A", "VEHICLE RECIEVING", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtIgnKey), "IGNKEY", "SMIS_MRRINV_TABLE"), "", "RR No:" & N2Str2Null(txtCode), "", ""
        '***************************************

        'LogAudit "A", "VEHICLE RECEVING", cboModelDescript & " CS: " & txtIgnKey & " DATE RECEVIED:" & txtDateReceived
    Else
        SQL_STATEMENT = "UPDATE SMIS_MRRINV_TABLE SET" & _
                      " CODE = " & vtxtCode & ", DESCRIPT = " & vtxtdescript & ", MAKE = " & vtxtmake & ", MODEL =" & vtxtmodel & "," & _
                      " YEER = " & vtxtyeer & ", CLASS = " & vcboclass & "," & _
                      " SOURCE = " & vcbosource & ", TRANSMISSION = " & vtxtTransmission & "," & _
                      " UNIT = " & vtxtunit & "," & _
                      " COLOR = " & vcbocolor & "," & _
                      " IGNKEY = " & vtxtignkey & ", MODELCODE = " & vtxtModelCode & ", " & _
                      " PRODNO = " & vtxtprodno & ", " & _
                      " SERIALNO = " & vtxtserialno & ", FRAMENO = " & vtxtframeno & ", " & _
                      " VINO = " & vtxtVINo & ", " & " PONO = " & vtxtpono & ", " & _
                      " ENGINENO = " & vtxtengineno & ", " & _
                      " FUELUSED = " & vtxtfuelused & ", " & _
                      " PISTONDISP = " & vtxtpistondisp & ", " & _
                      " GVW = " & vtxtgvw & ", " & _
                      " PURCHPRICE = " & vtxtpurchprice & ", MMPCSUBS = " & vtxtmmpcsubs & "," & _
                      " FBBODY = " & vtxtfbbody & ", " & _
                      " AIRCON = " & vtxtaircon & ", " & _
                      " STEREO = " & vtxtstereo & ", " & _
                      " CODEALARM = " & vtxtcodealarm & ", " & _
                      " PULLOUT = " & vtxtpullout & ", " & _
                      " LTO = " & vtxtlto & ", REFPONO = " & vtxtrefpono & ", DRNO= " & vtxtdrno & ", " & _
                      " TINT = " & vtxttint & ", REMARKS1 = " & vtxtremarks1 & ", REMARKS2 = " & vtxtremarks2 & ", REMARKS3 = " & vtxtremarks3 & ", LTOSTATUS = " & vtxtltostatus & ", CSR = " & vtxtcsr & ", NOTES = " & vtxtnote & ", ONSHOWROOM = " & voptonshowroom & ", WITHPROSBUYERS = " & voptwithprosbuyers & "," & _
                      " SEATCOVER = " & vtxtseatcover & ", MSPLUS = " & vtxtmspluscard & ", FLOORMAT =" & vtxtfloormat & ", " & _
                      " DATERECEIVED = " & vtxtdatereceived & ", PROFILE1 =" & vtxtprofile1 & ", PROFILE2 =" & vtxtprofile2 & ", PROFILE3 =" & vtxtprofile3 & ", PROFILE4 =" & vtxtprofile4 & ", PULLOUTDATE = " & vtxtpulloutdate & _
                      " WHERE ID = " & labid.Caption


        gconDMIS.Execute (SQL_STATEMENT)
        '****NEW LOG AUDIT*******

        NEW_LogAudit "E", "VEHICLE RECIEVING", SQL_STATEMENT, Null2String(labid), "", "RR No:" & Null2String(vtxtCode), "", ""
        '***********************

        'LogAudit "E", "VEHICLE RECEVING", cboModelDescript & " CS: " & txtIgnKey & " DATE RECEVIED:" & txtDateReceived
    End If
    'NOT REQUIRED FEATURE
    If Len(txtPO) > 0 And IsDate(txtDateReceived) = True Then
        gconDMIS.Execute "update smis_po set DateReceived=" & N2Date2Null(txtDateReceived) & " Where PO_NO=" & N2Str2Null(txtPO)

        '        Dim rsPO                        As ADODB.Recordset
        '        Set rsPO = gconDMIS.Execute("Select CUSCDE from SMIS_PO where PO_No=" & N2Str2Null(txtPO))
        '        If AddorEdit = "ADD" Then
        '            If Not rsPO.EOF Or Not rsPO.BOF Then
        '                If IsNull(rsPO!CUSCDE) = False Then
        '                    gconDMIS.Execute "update SMIS_MRRINV SET CUSTOMERCODE='" & rsPO!CUSCDE & "' where PONO=" & N2Str2Null(txtPO)
        '                    gconDMIS.Execute "update SMIS_MRRINV SET ISTATUS='A' WHERE PONO=" & N2Str2Null(txtPO) & " AND (ISTATUS='O') "
        '                End If
        '            End If
        '        Else
        '            If Null2String(rsMRRINV!ISTATUS) <> "A" Then
        '                If Not rsPO.EOF Or Not rsPO.BOF Then
        '                    If IsNull(rsPO!CUSCDE) = False Then
        '                        gconDMIS.Execute "update SMIS_MRRINV SET CUSTOMERCODE='" & rsPO!CUSCDE & "' where PONO=" & N2Str2Null(txtPO)
        '                        gconDMIS.Execute "update SMIS_MRRINV SET ISTATUS='A' WHERE PONO=" & N2Str2Null(txtPO) & " AND (ISTATUS='O') "
        '                    End If
        '                End If
        '            End If
        '        End If'
    End If
    If vtxtdatereleased <> "NULL" Then
        gconDMIS.Execute "update SMIS_MrrInv set datereleased = " & vtxtdatereleased & " WHERE IGNKEY = " & vtxtignkey
    Else
        gconDMIS.Execute "update SMIS_MrrInv set datereleased = " & vtxtdatereleased & " WHERE IGNKEY = " & vtxtignkey
    End If

    '*****RESET THE SQL_STATEMENT VARIABLE********
    SQL_STATEMENT = ""
    '********************************************

    'TO BE ABLE TO EDIT VEHICLES DETAILS EVEN INVOICED
    If AddorEdit = "EDIT" Then
        Dim SQL
        SQL = "UPDATE SMIS_SALESORDER SET " & vbCrLf
        SQL = SQL & " MODEL = " & N2Str2Null(txtModel) & "," & vbCrLf
        SQL = SQL & " PRODNO = " & N2Str2Null(txtProdNo) & "," & vbCrLf
        SQL = SQL & " MODELDESCRIPTION= " & N2Str2Null(cboModelDescript) & "," & vbCrLf
        SQL = SQL & " ENGINENO = " & N2Str2Null(txtEngineNo) & "," & vbCrLf
        SQL = SQL & " IGNKEY_NO = " & N2Str2Null(txtIgnKey) & "," & vbCrLf
        SQL = SQL & " FRAMENO = " & N2Str2Null(txtFrameNo) & "," & vbCrLf
        SQL = SQL & " VINO = " & N2Str2Null(txtVINO) & "," & vbCrLf
        SQL = SQL & " COLOR = " & N2Str2Null(cboColor) & vbCrLf
        SQL = SQL & " WHERE ignkey_no= '" & rsMRRINV!ignkey & "'"
        gconDMIS.Execute SQL

        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "SALES ORDER", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtIgnKey), "IGNKEY_NO", "SMIS_SALESORDER"), "", "RR No:" & txtCode, "", ""
    End If

    rsRefresh
    rsMRRINV.Find "IGNKEY =" & vtxtignkey

    cmdCancel.Value = True
    Exit Sub
End Sub

Private Sub cmdServiceVehicle_Click()
    txtVehDet = ""
    txtVehDet = txtVehDet & " REGISTRATION TO :" & (txtVeh_register_CustCode) & vbCrLf
    txtVehDet = txtVehDet & " CUSTOMER NAME   :" & (txtVeh_register_CustName) & vbCrLf
    txtVehDet = txtVehDet & " VIN#            :" & (txtVINO) & vbCrLf
    txtVehDet = txtVehDet & " PLATE#          :" & (txtIgnKey) & vbCrLf
    txtVehDet = txtVehDet & " CS #            :" & (txtIgnKey) & vbCrLf
    txtVehDet = txtVehDet & " YEAR            :" & (txtYeer) & vbCrLf
    txtVehDet = txtVehDet & " MAKE            :" & (TXTMAKE) & vbCrLf
    txtVehDet = txtVehDet & " MODEL           :" & (txtModel) & vbCrLf
    txtVehDet = txtVehDet & " MODEL CODE      :" & (txtModelCode) & vbCrLf
    txtVehDet = txtVehDet & " ENGINE #        :" & (txtEngineNo) & vbCrLf
    txtVehDet = txtVehDet & " PRODUCTION #    :" & (txtProdNo) & vbCrLf
    txtVehDet = txtVehDet & " SERIAL#         :" & (txtSerialNo) & vbCrLf
    txtVehDet = txtVehDet & " DESCRIPTION     :" & (cboModelDescript) & vbCrLf
    txtVehDet = txtVehDet & " COLOR #         :" & (SetColor(cboColor)) & vbCrLf
    txtVehDet = txtVehDet & " SELLING_DEALER  :" & COMPANY_CODE & vbCrLf
    txtVehDet = txtVehDet & " MAKE            :Mitsubishi"


    ShowHidePictureBox2 picServiceInfo, True, picBottoms

End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_unPost", "VEHICLE RECIEVING") = False Then Exit Sub
    On Error GoTo ErrorCode:
    
    If MsgBox("Are You Sure You Want To Un-Post This Transaction", vbInformation + vbYesNo) = vbNo Then: Exit Sub
    SQL_STATEMENT = "update SMIS_MrrInv_Table set status='U' where id = " & labid.Caption
    gconDMIS.Execute (SQL_STATEMENT)
    '**********NEW LOG AUDIT***********
        NEW_LogAudit "U", "VEHICLE RECIEVING", SQL_STATEMENT, Null2String(labid), "", "RR No:" & txtCode, "", ""
    '*********************************

    MessagePop RecSaveOk, "Transaction Unposted", "Record Sucessfully Un-Posted", 1000, 2
    
    'LogAudit "U", "VEHICLE RECEVING", cboModelDescript & " CS: " & txtIgnKey & " DATE RECEVIED:" & txtDateReceived
    rsRefresh
    rsMRRINV.Find ("ID=" & labid)
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdRegisterVehicleClose_Click()
    ShowHidePictureBox2 picServiceInfo, False, picBottoms
End Sub

Private Sub Command3_Click()
    SearchMaster.SearchForMRR
    SearchMaster.Show 1
End Sub

Private Sub Command4_Click()
    '    If AddorEdit = "EDIT" Then
    If Function_Access(LOGID, "ACESS_SYSTEM", "VEHICLE RECIEVING") = False Then Exit Sub
    txtPullOutDate.Enabled = True: txtPullOutDate.SetFocus
    '   End If
End Sub

Private Sub Command5_Click()
    '    If AddorEdit = "EDIT" Then
    If Function_Access(LOGID, "ACESS_SYSTEM", "VEHICLE RECIEVING") = False Then Exit Sub
    txtDateReceived.Enabled = True: txtDateReceived.SetFocus
    '   End If
End Sub

Private Sub cmdRegisterVehicle_Click()
    Dim ColorCode                                                     As String
    Dim rsCusVeh1                                                     As ADODB.Recordset
    Dim SQL                                                           As String
    Dim rsHanapID                                                     As ADODB.Recordset
    Dim vID                                                           As String
    Dim rsVindup                                                      As ADODB.Recordset
    Dim rsPlatedup                                                      As ADODB.Recordset
    
    On Error GoTo next_BABY
    If TXTMAKE.Text = "" Then
        ShowIsRequiredMsg ("Make Cannot be Blank")
        Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------------
    'updated by:    IEBV 11032010_1004Am
    'description:   Vin No and plate number cannot be duplicate
    
'    If AddorEdit = "EDIT" Then
'        If rsMRRINV!ignkey <> txtIgnKey.Text Then
'            Dim rsdunignkey As ADODB.Recordset
'            Set rsdunignkey = New ADODB.Recordset
'            Set rsdunignkey = gconDMIS.Execute("Select * from csms_cusveh where Plate_no = " & N2Str2Null(txtIgnKey) & "")
'            If Not rsdunignkey.EOF And Not rsdunignkey.BOF Then
'                On Error Resume Next
'                MsgBox N2Str2Null(txtIgnKey) & "is aready" & vbCrLf & "registered to " & rsdunignkey!NIYM & ". ", vbInformation + vbOKOnly
'                txtIgnKey.SetFocus
'                Set rsPlatedup = Nothing
'                Exit Sub
'            End If
'        End If
'    Else
'        Set rsVindup = New ADODB.Recordset
'        Set rsVindup = gconDMIS.Execute("Select * from csms_cusveh where vin = " & N2Str2Null(txtVINO) & "")
'        If Not rsVindup.EOF And Not rsVindup.BOF Then
'            On Error Resume Next
'            MsgBox N2Str2Null(txtVINO) & " is already " & vbCrLf & "registered to " & rsVindup!NIYM & ". ", vbInformation + vbOKOnly
'            txtVINO.SetFocus
'            Set rsVindup = Nothing
'            Exit Sub
'        End If
'
'        Set rsPlatedup = New ADODB.Recordset
'        Set rsPlatedup = gconDMIS.Execute("Select * from csms_cusveh where Plate_no = " & N2Str2Null(txtIgnKey) & "")
'        If Not rsPlatedup.EOF And Not rsPlatedup.BOF Then
'            On Error Resume Next
'            MsgBox N2Str2Null(txtIgnKey) & "is aready" & vbCrLf & "registered to " & rsPlatedup!NIYM & ". ", vbInformation + vbOKOnly
'            txtIgnKey.SetFocus
'            Set rsPlatedup = Nothing
'            Exit Sub
'        End If
'    End If
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    ColorCode = SetColor(cboColor)
    Set rsCusVeh1 = New ADODB.Recordset
    'Set rsCusVeh1 = gconDMIS.Execute("Select * From CSMS_CUSVEH WHERE UPPER(MAKE) = 'HYUNDAI' AND VCOND_NO='" & txtIgnKey & "' AND INVOICENO IS NULL")
    Set rsCusVeh1 = gconDMIS.Execute("Select * From CSMS_CUSVEH WHERE UPPER(MAKE) = 'MITSUBISHI' AND VCOND_NO='" & txtIgnKey & "'")

    If Not (rsCusVeh1.EOF Or rsCusVeh1.BOF) Then
        SQL = " Update CSMS_CUSVEH SET  "
        SQL = SQL & " CUSCDE=" & N2Str2Null(txtVeh_register_CustCode) & ", "
        SQL = SQL & " NIYM= " & N2Str2Null(txtVeh_register_CustName) & ", "
        SQL = SQL & " VIN=" & N2Str2Null(txtVINO) & ", "
        SQL = SQL & " PLATE_NO= " & N2Str2Null(txtIgnKey) & ", "
        SQL = SQL & " VCOND_NO= " & N2Str2Null(txtIgnKey) & ", "
        SQL = SQL & " YER= " & N2Str2Null(txtYeer) & ", "
        SQL = SQL & " MAKE= " & N2Str2Null(TXTMAKE) & ", "
        SQL = SQL & " MODEL= " & N2Str2Null(txtModel) & ", "
        SQL = SQL & " MODELCODE= " & N2Str2Null(txtModelCode) & ", "
        SQL = SQL & " ENGINE= " & N2Str2Null(txtEngineNo) & ", "
        SQL = SQL & " PRODNO= " & N2Str2Null(txtProdNo) & ", "
        SQL = SQL & " SERIAL= " & N2Str2Null(txtSerialNo) & ", "
        SQL = SQL & " DESCRIPTION= " & N2Str2Null(cboModelDescript) & ", "
        SQL = SQL & " CLRCDE= " & N2Str2Null(ColorCode) & ", "
        SQL = SQL & " SELLING_DEALER='" & COMPANY_CODE & "'" & "  "
        SQL = SQL & " WHERE UPPER(MAKE) = 'Mitsubishi' AND VCOND_NO='" & txtIgnKey & "'"

        gconDMIS.Execute SQL


        '*********UPDATED BY RDC Aug 28 2008
        Set rsHanapID = gconDMIS.Execute("SELECT * FROM CSMS_CUSVEH WHERE VCOND_NO='" & Null2String(txtIgnKey) & "'")

        If Not (rsHanapID.EOF And rsHanapID.BOF) Then
            vID = Null2String(rsHanapID!ID)
        End If
        '************NEW LOG AUDIT**************
        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "CUSTOMER VEHICLE", SQL_STATEMENT, Null2String(vID), "", "RR No:" & txtCode, "", ""
        '**************************************
        MessagePop RecSaveOk, "Service Vehicle Registration", "Vehicle Registration Updated"
        'LogAudit "E", "NEW VEHICLE REGISTRATION", "CONDUCTION STRICKER" & txtIgnKey & " MODEL & txtModel " & " VIN" & txtVINO
    Else
        SQL = " INSERT INTO CSMS_CUSVEH  ( CUSCDE, NIYM, VIN, PLATE_NO, VCOND_NO,YER, MAKE, "
        SQL = SQL & " MODEL, MODELCODE,ENGINE,PRODNO, SERIAL, DESCRIPTION,CLRCDE ,SELLING_DEALER ) VALUES ( "
        SQL = SQL & N2Str2Null(txtVeh_register_CustCode) & " ,"
        SQL = SQL & N2Str2Null(txtVeh_register_CustName) & " ,"
        SQL = SQL & N2Str2Null(txtVINO) & " ,"
        SQL = SQL & N2Str2Null(txtIgnKey) & " ,"
        SQL = SQL & N2Str2Null(txtIgnKey) & " ,"
        SQL = SQL & N2Str2Null(txtYeer) & " ,"
        SQL = SQL & N2Str2Null(TXTMAKE) & " ,"
        SQL = SQL & N2Str2Null(txtModel) & " ,"
        SQL = SQL & N2Str2Null(txtModelCode) & " ,"
        SQL = SQL & N2Str2Null(txtEngineNo) & " ,"
        SQL = SQL & N2Str2Null(txtProdNo) & " ,"
        SQL = SQL & N2Str2Null(txtSerialNo) & " ,"
        SQL = SQL & N2Str2Null(cboModelDescript) & " ,"
        SQL = SQL & N2Str2Null(ColorCode) & " ,"
        SQL = SQL & N2Str2Null(COMPANY_CODE) & ")"
        gconDMIS.Execute SQL
        SQL_STATEMENT = SQL
        Set rsHanapID = gconDMIS.Execute("SELECT * FROM CSMS_CUSVEH WHERE VCOND_NO='" & Null2String(txtIgnKey) & "'")

        If Not (rsHanapID.EOF And rsHanapID.BOF) Then
            vID = Null2String(rsHanapID!ID)
        End If
        NEW_LogAudit "A", "CUSTOMER VEHICLE", SQL_STATEMENT, Null2String(vID), "", "RR No:" & txtCode, "", ""
        '**************************************
        MessagePop RecSaveInfo, "Service Vehicle Registration", "New Vehicle Registration Sucessful"
        'LogAudit "A", "NEW VEHICLE REGISTRATION", "CONDUCTION STRICKER" & txtIgnKey & " MODEL & txtModel " & " VIN" & txtVINO
    End If

    cmdRegisterVehicleClose_Click
    rsRefresh
    rsMRRINV.Find ("ID=" & labid)
    StoreMemVars
    
    Exit Sub
    
next_BABY:
    Dim RSTMP As New ADODB.Recordset
    
    Set RSTMP = gconDMIS.Execute("SELECT NIYM, VIN FROM CSMS_CUSVEH WHERE VIN = " & N2Str2Null(txtVINO) & "")
    If (RSTMP.BOF And RSTMP.EOF) Then
        Set RSTMP = New ADODB.Recordset
        Set RSTMP = gconDMIS.Execute("SELECT NIYM, PLATE_NO FROM CSMS_CUSVEH WHERE PLATE_NO = " & N2Str2Null(txtIgnKey) & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            MsgBox "Plate no already exist in Vehicle master file and register to : " & Null2String(RSTMP!NIYM) & ""
        Else
            MsgBox Err.Description
        End If
    Else
        MsgBox "Vin no already exist in Vehicle master file and register to : " & Null2String(RSTMP!NIYM) & ""
    End If
    
    Err.Clear
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE RECIEVING)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labid), "VEHICLE RECIEVING")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Set SearchMaster = New frmSMIS_Mis_SearchMaster

    rsRefresh
    If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
        rsMRRINV.MoveLast
    End If
    initMemvars
    InitData

    AddorEdit = ""
    picTops.Enabled = False
    picVehicleReceving.Enabled = False
    picSaves.Visible = False
    picAdds.Visible = True
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub rsRefresh()
    Set rsMRRINV = New ADODB.Recordset
    rsMRRINV.CursorLocation = adUseClient
    Call rsMRRINV.Open("SELECT * from SMIS_MrrInv_Table order by id ASC", gconDMIS, adOpenKeyset)
End Sub

Private Sub SearchMaster_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    '
    If XSelection = "MRR" Then
        txtNote = Null2String(oCusRs!Notes)
        txtProfile1 = Null2String(oCusRs!profile1)
        txtProfile2 = Null2String(oCusRs!profile2)
        txtProfile3 = Null2String(oCusRs!profile3)
        txtProfile4 = Null2String(oCusRs!profile4)
    Else

        txtPO = Null2String(oCusRs!po_no)
        cboModelDescript = Null2String(oCusRs!ModelDescript)
        TXTMAKE = "Mitsubishi"
        txtModel = Null2String(oCusRs!Model)
        txtModelCode = Null2String(oCusRs!ModelCode)
        SetClass
        txtYeer = Null2String(oCusRs!MODELYEAR)
        cboSource = Null2String(oCusRs!Source)
        txtUnit = Null2String(oCusRs!ModelDescript)
        cboColor = Null2String(oCusRs!Color)
        txtFuelUsed = Null2String(oCusRs!Fuel)
        txtNote = Null2String(oCusRs!Notes)
        txtPurchPrice = NumericVal(oCusRs!CD_AMOUNT)
        txtSubisidy = NumericVal(oCusRs!SUBSIDY)
        optWithProsBuyers.Value = True
        txtRemarks1 = Null2String(oCusRs!Notes)
    End If
    Unload SearchMaster

End Sub

Private Sub StoreMemVars()
    If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then

        labid.Caption = rsMRRINV!ID
        labAPJ = CheckAPJNum(Null2String(rsMRRINV!CODE), "VEHICLES"): labDetails = ""

        '†-------------------------------------------------------------
        'UPDATE BY   : MJP 08192008 04:18 PM
        'DESCRIPTION :
        'labAPJ.Caption = ReturnAccountsPayableInformation(Null2String(rsMRRINV!ignkey), Null2String(rsMRRINV!CODE))(0)
        'labDetails.Caption = ReturnAccountsPayableInformation(Null2String(rsMRRINV!ignkey), Null2String(rsMRRINV!CODE))(1)
        'UPDATE BY   : MJP 08192008 04:18 PM
        '-------------------------------------------------------------†

        LABALLOWREPRINT = Null2String(rsMRRINV!PRINTED)
        txtCode = Null2String(rsMRRINV!CODE)

        txtPO = Null2String(rsMRRINV!PONO)
        cboModelDescript = Null2String(LTrim(RTrim(rsMRRINV!DESCRIPT)))
        TXTMAKE = Null2String(rsMRRINV!Make)
        txtModel = Null2String(rsMRRINV!Model)
        txtModelCode = Null2String(rsMRRINV!ModelCode)
        SetClass
        txtYeer = Null2String(rsMRRINV!YEER)
        cboSource.Text = LTrim(RTrim(Null2String(rsMRRINV!Source)))
        txtUnit = Null2String(rsMRRINV!unit)
        cboColor = Null2String(rsMRRINV!Color)
        txtIgnKey = Null2String(rsMRRINV!ignkey)
        txtProdNo = Null2String(rsMRRINV!prodno)
        cboTransmission = Null2String(rsMRRINV!Transmission)
        txtSerialNo = Null2String(rsMRRINV!SERIALNO)
        txtVINO = Null2String(rsMRRINV!VINO)
        txtEngineNo = Null2String(rsMRRINV!EngineNo)
        txtFuelUsed = Null2String(rsMRRINV!fuelused)
        txtPistonDisp = Null2String(rsMRRINV!pistondisp)
        txtFrameNo = Null2String(rsMRRINV!frameno)
        txtGVW = Null2String(rsMRRINV!gvw)

        txtPurchPrice = NumericVal(rsMRRINV!PurchPrice)
        txtSubisidy = NumericVal(rsMRRINV!MMPCSUBs)
        txtFBBody = NumericVal(rsMRRINV!fbbody)
        txtAircon = NumericVal(rsMRRINV!aircon)
        txtStereo = NumericVal(rsMRRINV!stereo)
        txtCodeAlarm = NumericVal(rsMRRINV!codealarm)
        txtPullOut = NumericVal(rsMRRINV!pullout)
        txtLto = NumericVal(rsMRRINV!LTO)
        txtTint = NumericVal(rsMRRINV!tint)
        txtSeatCover = NumericVal(rsMRRINV!seatcover)
        txtMSPlusCard = NumericVal(rsMRRINV!msplus)
        txtFloormat = NumericVal(rsMRRINV!floormat)

        txtref_PONO = Null2String(rsMRRINV!refPONO)

        txtPullOutDate = Null2String(rsMRRINV!PullOutDate)

        txtDateReceived = Null2String(rsMRRINV!datereceived)
        txtDateReleased = Null2String(rsMRRINV!DateReleased)

        txtProfile1 = Null2String(rsMRRINV!profile1)
        txtProfile2 = Null2String(rsMRRINV!profile2)
        txtProfile3 = Null2String(rsMRRINV!profile3)
        txtProfile4 = Null2String(rsMRRINV!profile4)
        txtRemarks1 = Null2String(rsMRRINV!Remarks1)
        txtRemarks2 = Null2String(rsMRRINV!Remarks2)
        txtRemarks3 = Null2String(rsMRRINV!Remarks3)
        txtLTOStatus = Null2String(rsMRRINV!LTOStatus)
        txtCSR = Null2String(rsMRRINV!CSR)
        txtNote = Null2String(rsMRRINV!Notes)

        ''STATUS INDICATOR LINE
        Dim RELEASEINFO, ISTATUS
        Dim RELEASED
        Dim rsInvStatus                                               As ADODB.Recordset
        Dim temprs                                                    As ADODB.Recordset

        RELEASEINFO = Null2String(rsMRRINV!RELEASED)
        labInventoryStatus = ""
        ISTATUS = Null2String(rsMRRINV!ISTATUS)
        'sold and unrealease



        If ISTATUS = "S" And RELEASEINFO = False Then
            Set rsInvStatus = gconDMIS.Execute("Select CUSNAM  from ALL_CUSMAS WHERE CUSCDE='" & rsMRRINV!CustomerCode & "'")
            If Not (rsInvStatus.EOF Or rsInvStatus.BOF) Then
                labInventoryStatus = "** INVOICED / NOT RELEASED **" & Null2String(rsInvStatus!CUSNAM)
            Else
                labInventoryStatus = "** INVOICED / NOT RELEASED ** CUSTOMER INFORMATION MISSING"
            End If
            '  picVehicleDetails.Enabled = False
            '  picModelDetails.Enabled = False
            '  picRefHeader.Enabled = False
            '  cmdDelete.Enabled = False
            'sold and released
        ElseIf ISTATUS = "R" And RELEASEINFO = True Then
            Set rsInvStatus = gconDMIS.Execute("Select CUSNAM  from ALL_CUSMAS WHERE CUSCDE='" & rsMRRINV!CustomerCode & "'")
            If Not (rsInvStatus.EOF And rsInvStatus.BOF) Then
                labInventoryStatus = "** SOLD  TO **" & Null2String(rsInvStatus!CUSNAM)
            Else
                labInventoryStatus = "** SOLD BUT CUSTOMER INFORMATION MISSING : Check Information in Customer Master File***"
            End If
            ' picVehicleDetails.Enabled = False
            ' picModelDetails.Enabled = False
            ' picRefHeader.Enabled = False
            ' cmdDelete.Enabled = False
            'allocated
        ElseIf ISTATUS = "A" Then
            Set rsInvStatus = gconDMIS.Execute("Select CUSNAM  from ALL_CUSMAS WHERE CUSCDE='" & rsMRRINV!CustomerCode & "'")
            If Not (rsInvStatus.EOF And rsInvStatus.BOF) Then
                labInventoryStatus = "** ALLOCATED FOR **" & Null2String(rsInvStatus!CUSNAM)
            Else
                labInventoryStatus = "**** ALLOCATED / CUSTOMER INFORMATION MISSING : Check Information in Customer Master File***"
            End If
            'picVehicleDetails.Enabled = True
            'picModelDetails.Enabled = True
            'picRefHeader.Enabled = True
            'cmdDelete.Enabled = False
        ElseIf ISTATUS = "T" Then
            Set temprs = gconDMIS.Execute("Select UPPER(Entity_to)  Entity_To From SMIS_StockTransfer where VSNO=" & N2Str2Null(rsMRRINV!prodno))
            If Not (temprs.EOF Or temprs.BOF) Then
                labInventoryStatus = "**UNIT TRANSFERED " & Null2String(temprs!ENTITY_TO) & " **"
            Else
                labInventoryStatus = "**UNIT TRANSFERED MISSING TRANSFEREE INFO**"
            End If
            'picVehicleDetails.Enabled = False
            'picModelDetails.Enabled = False
            'picRefHeader.Enabled = False
            'cmdDelete.Enabled = False
        ElseIf ISTATUS = "O" Or ISTATUS = "" Then
            picRefHeader.Enabled = True
            picVehicleDetails.Enabled = True
            picModelDetails.Enabled = True
            labInventoryStatus = "** AVAILABLE/OPEN**"
            cmdDelete.Enabled = True
        End If

        If Null2String(rsMRRINV!OnShowroom) = "Y" Then
            optOnShowroom.Value = True
        Else
            optOnShowroom.Value = False
        End If
        If Null2String(rsMRRINV!WithProsBuyers) = "Y" Then
            optWithProsBuyers.Value = True
        Else
            optWithProsBuyers.Value = False
        End If

        If Null2String(rsMRRINV!STATUS) = "P" Then
            labStatus = "**POSTED**"
            cmdPrint.Enabled = True: cmdPost.Enabled = False: cmdUnPost.Enabled = True: cmdCancelCO.Enabled = False: cmdEdit.Enabled = False: cmdDelete.Enabled = False
            If labAPJ <> "" Then
                labDetails = "IMPORTED TO ACCOUNTING"
                cmdPost.Enabled = False
                cmdUnPost.Enabled = False
                cmdPrint.Enabled = False
                cmdCancelCO.Enabled = False
                cmdEdit.Enabled = False
            End If
        ElseIf Null2String(rsMRRINV!STATUS) = "U" Then
            labStatus = ""
            cmdPrint.Enabled = False: cmdPost.Enabled = True: cmdUnPost.Enabled = False: cmdEdit.Enabled = True
            ' If ISTATUS = "O" And IsDate(RELEASEINFO) = False Then
            cmdDelete.Enabled = True
            cmdCancelCO.Enabled = True
            cmdEdit.Enabled = True
            'Else
            '   cmdDelete.Enabled = False
            '  cmdCancelCO.Enabled = False
            ' cmdEdit.Enabled = False
            'End If
        ElseIf Null2String(rsMRRINV!STATUS) = "" Then
            labStatus = ""
            cmdPrint.Enabled = False: cmdPost.Enabled = True: cmdUnPost.Enabled = False

            'If ISTATUS = "O" And IsDate(RELEASEINFO) = False Then
            cmdDelete.Enabled = True
            cmdEdit.Enabled = True
            cmdCancelCO.Enabled = True
            'Else
            '   cmdDelete.Enabled = False
            '   cmdCancelCO.Enabled = False
            '   cmdEdit.Enabled = False
            'End If
        ElseIf Null2String(rsMRRINV!STATUS) = "C" Then
            labStatus = "**CANCELLED**"
            labInventoryStatus = ""
            cmdPrint.Enabled = False: cmdPost.Enabled = False: cmdUnPost.Enabled = False: cmdCancelCO.Enabled = False: cmdEdit.Enabled = False: cmdDelete.Enabled = False
        End If


        txtVeh_register_CustCode = ""
        txtVeh_register_CustName = ""

        If Null2String(rsMRRINV!CustomerCode) <> "" Then
            txtVeh_register_CustCode = rsMRRINV!CustomerCode
        End If


        Dim rsCusVeh1                                                 As ADODB.Recordset
        Set rsCusVeh1 = gconDMIS.Execute("Select CUSCDE,NIYM,INVOICENO From CSMS_CUSVEH WHERE UPPER(MAKE) = 'Mitsubishi' AND VCOND_NO='" & txtIgnKey & "'")
        If Not (rsCusVeh1.EOF Or rsCusVeh1.BOF) Then
            LABVEHREGISTER = "VEHICLE REGISTER:" & Null2String(rsCusVeh1!NIYM)
        Else
            LABVEHREGISTER = "NO VEHICLE REGISTRATION TO SERVICE"
        End If

        '        If Null2String(rsMRRINV!status) = "P" Then
        '            cmdServiceVehicle.Enabled = True
        '        Else
        '            cmdServiceVehicle.Enabled = False
        '        End If
        '

        txtref_PONO = Null2String(rsMRRINV!refPONO)
        txtDRNO = Null2String(rsMRRINV!drno)
        TotalCost
        txtTotalCost = Tutal

    Else
        ShowNoRecord
        cmdAdd.Value = True

    End If
End Sub

Private Sub txtVeh_register_CustCode_Change()
    txtVeh_register_CustName = ""
    cmdRegisterVehicle.Enabled = False
    If Len(txtVeh_register_CustCode) >= 6 Then
        Dim rsCust                                                    As ADODB.Recordset
        Set rsCust = gconDMIS.Execute("select ACCTNAME from all_customer_table where cuscde=" & N2Str2Null(Repleys(txtVeh_register_CustCode)))
        If Not (rsCust.EOF Or rsCust.BOF) Then
            txtVeh_register_CustName = Null2String(rsCust!AcctName)
            cmdRegisterVehicle.Enabled = True
        End If

        txtVehDet = ""
        txtVehDet = txtVehDet & " CODE           :" & (txtVeh_register_CustCode) & vbCrLf
        txtVehDet = txtVehDet & " NAME           :" & (txtVeh_register_CustName) & vbCrLf
        txtVehDet = txtVehDet & " VIN#           :" & (txtVINO) & vbCrLf
        txtVehDet = txtVehDet & " PLATE#         :" & (txtIgnKey) & vbCrLf
        txtVehDet = txtVehDet & " CS #           :" & (txtIgnKey) & vbCrLf
        txtVehDet = txtVehDet & " YEAR           :" & (txtYeer) & vbCrLf
        txtVehDet = txtVehDet & " MAKE           :" & (TXTMAKE) & vbCrLf
        txtVehDet = txtVehDet & " MODEL          :" & (txtModel) & vbCrLf
        txtVehDet = txtVehDet & " MODEL CODE     :" & (txtModelCode) & vbCrLf
        txtVehDet = txtVehDet & " ENGINE #       :" & (txtEngineNo) & vbCrLf
        txtVehDet = txtVehDet & " PRODUCTION #   :" & (txtProdNo) & vbCrLf
        txtVehDet = txtVehDet & " SERIAL#        :" & (txtSerialNo) & vbCrLf
        txtVehDet = txtVehDet & " DESCRIPTION    :" & (cboModelDescript) & vbCrLf
        txtVehDet = txtVehDet & " COLOR #        :" & (SetColor(cboColor)) & vbCrLf
        txtVehDet = txtVehDet & " SELLING DEALER :" & COMPANY_CODE & vbCrLf
        txtVehDet = txtVehDet & " MAKE           :Mitsubishi"

        Set rsCust = Nothing
    End If

End Sub

Private Sub Timer2_Timer()

    If labInventoryStatus.Caption <> "" Then
        If labInventoryStatus.Visible = True Then
            labInventoryStatus.Visible = False
        Else
            labInventoryStatus.Visible = True
        End If
    End If
    If labStatus <> "" Then
        If labStatus.Visible = True Then
            labStatus.Visible = False
        Else
            labStatus.Visible = True
        End If
    End If

End Sub

Private Sub TotalCost()
    On Error GoTo ADDER:
    Dim pp, mmpc, fb, ai, st, co, pu, lt, ti, se, ms, fl, acc         As Currency
    pp = NumericVal(txtPurchPrice)
    mmpc = NumericVal(txtSubisidy)
    fb = NumericVal(txtFBBody)
    ai = NumericVal(txtAircon)
    st = NumericVal(txtStereo)
    co = NumericVal(txtCodeAlarm)
    pu = NumericVal(txtPullOut)
    lt = NumericVal(txtLto)
    ti = NumericVal(txtTint)
    se = NumericVal(txtSeatCover)
    ms = NumericVal(txtMSPlusCard)
    fl = NumericVal(txtFloormat)
    acc = NumericVal(txtAccSummation)
    Tutal = (pp - mmpc) + fb + ai + st + co + pu + lt + ti + se + ms + fl + acc

    txtSubTotalCost = (pp - mmpc) + acc
    Exit Sub
ADDER:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub txtAccSummation_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtAccSummation_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtAircon_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtAircon_GotFocus()
    If NumericVal(txtAircon) <= 0 Then txtAircon = ""
End Sub

Private Sub txtAircon_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtAircon_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtAircon_LostFocus()
    If NumericVal(txtAircon) <= 0 Then txtAircon = "0.00"
End Sub

Private Sub txtCode_LostFocus()

    If AddorEdit = "" Then Exit Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    txtCode = Format(txtCode, "000000")
    Dim lngcount                                                      As Integer
    lngcount = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_MrrInv_table WHERE upper(LTrim(RTrim(code)))=" & N2Str2Null(UCase(LTrim(RTrim(txtCode))))).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lngcount >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "MRR Number Already Exist"
            txtCode.SetFocus
            Exit Sub
        End If
    Else
        If lngcount >= 1 And Null2String(UCase(LTrim(RTrim(rsMRRINV!CODE)))) <> UCase(LTrim(RTrim(txtCode))) Then
            MessagePop RecSaveWarning, "Duplicate Record", "MRR Number Already Exist"
            txtCode.SetFocus
            Exit Sub
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub txtCodeAlarm_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtCodeAlarm_GotFocus()
    If NumericVal(txtCodeAlarm) <= 0 Then txtCodeAlarm = ""
End Sub

Private Sub txtCodeAlarm_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtCodeAlarm_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCodeAlarm_LostFocus()
    If NumericVal(txtCodeAlarm) <= 0 Then txtCodeAlarm = "0.00"
End Sub

Private Sub txtDateReceived_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtDateReleased_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub
'Commented by: JUN 10232008
'Private Sub txtEngineNo_Change()
'    'UPDATED BY: JUN CEDRON
'    'DATE UPDATED: 08/13/2008
'    'DESCRIPTION: AVOID DUPLICATE FIELD
'
'    If AddorEdit = "ADD" Or AddorEdit = "EDIT" Then
'        txtFrameNo = txtEngineNo
'    End If
'End Sub

Private Sub txtEngineNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtFBBody_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtFBBody_GotFocus()
    If NumericVal(txtFBBody) <= 0 Then txtFBBody = ""
End Sub

Private Sub txtFBBody_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtFBBody_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFBBody_LostFocus()
    If NumericVal(txtFBBody) <= 0 Then txtFBBody = "0.00"
End Sub

Private Sub txtFloormat_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtFloormat_GotFocus()
    If NumericVal(txtFloormat) <= 0 Then txtFloormat = ""
End Sub

Private Sub txtFloormat_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtFloormat_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtFloormat_LostFocus()
    If NumericVal(txtFloormat) <= 0 Then txtFloormat = "0.00"
End Sub

Private Sub txtFuelUsed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtIgnKey_Change()
    '    If AddorEdit = "ADD" Then: txtcode = txtIgnKey
End Sub

Private Sub txtIgnKey_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtLto_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtLto_GotFocus()
    If NumericVal(txtLto) <= 0 Then txtLto = ""
End Sub

Private Sub txtLto_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtLto_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtLto_LostFocus()
    If NumericVal(txtLto) <= 0 Then txtLto = "0.00"
End Sub

Private Sub txtSubisidy_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtSubisidy_GotFocus()
    If NumericVal(txtSubisidy) <= 0 Then txtSubisidy = ""
End Sub

Private Sub txtSubisidy_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtSubisidy_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtSubisidy_LostFocus()
    If NumericVal(txtSubisidy) <= 0 Then txtSubisidy = "0.00"
End Sub

Private Sub txtMSPlusCard_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtMSPlusCard_GotFocus()
    If NumericVal(txtMSPlusCard) <= 0 Then txtMSPlusCard = ""
End Sub

Private Sub txtMSPlusCard_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtMSPlusCard_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtMSPlusCard_LostFocus()
    If NumericVal(txtMSPlusCard) <= 0 Then txtMSPlusCard = "0.00"
End Sub

Private Sub txtPistonDisp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtProdNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPullOut_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtPullOut_GotFocus()
    If NumericVal(txtPullOut) <= 0 Then txtPullOut = ""
End Sub

Private Sub txtPullOut_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtPullOut_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtPullOut_LostFocus()
    If NumericVal(txtPullOut) <= 0 Then txtPullOut = "0.00"
End Sub

Private Sub txtPullOutDate_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtPullOutDate_LostFocus()
    If IsDate(txtPullOutDate) = True Then
        txtDateReceived = txtPullOutDate
    End If
End Sub

Private Sub txtPurchPrice_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtPurchPrice_GotFocus()
    If NumericVal(txtPurchPrice) <= 0 Then txtPurchPrice = ""
End Sub

Private Sub txtPurchPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtPurchPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtPurchPrice_LostFocus()
    If NumericVal(txtPurchPrice) <= 0 Then txtPurchPrice = "0.00"
End Sub

Private Sub txtSeatCover_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtSeatCover_GotFocus()
    If NumericVal(txtSeatCover) <= 0 Then txtSeatCover = ""
End Sub

Private Sub txtSeatCover_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtSeatCover_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtSeatCover_LostFocus()
    If NumericVal(txtSeatCover) <= 0 Then txtSeatCover = "0.00"
End Sub

Private Sub txtSerialNo_Change()
    '    If AddorEdit = "ADD" And COMPANY_CODE = "HBK" Then
    '        txtVINo = txtSerialNo
    '        txtFrameNo = txtSerialNo
    '    ElseIf AddorEdit = "ADD" And COMPANY_CODE = "HGC" Then
    '        txtVINo = txtSerialNo
    '    End If
    'UPDATED BY: JUN CEDRON
    'DATE UPDATED: 08/13/2008

    If AddorEdit = "ADD" Or AddorEdit = "EDIT" Then
        txtVINO = txtSerialNo
        txtFrameNo = txtSerialNo
    End If

End Sub

Private Sub txtSerialNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtStereo_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtStereo_GotFocus()
    If NumericVal(txtStereo) <= 0 Then txtStereo = ""
End Sub

Private Sub txtStereo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtStereo_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtStereo_LostFocus()
    If NumericVal(txtStereo) <= 0 Then txtStereo = "0.00"
End Sub

Private Sub txtTint_Change()
    If AddorEdit = "" Then Exit Sub
    TotalCost
    txtTotalCost = Tutal
End Sub

Private Sub txtTint_GotFocus()
    If NumericVal(txtTint) <= 0 Then txtTint = ""
End Sub

Private Sub txtTint_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtTint_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtTint_LostFocus()
    If NumericVal(txtTint) <= 0 Then txtTint = "0.00"
End Sub

Private Sub txtTotalCost_GotFocus()
    If AddorEdit = "" Then Exit Sub
    txtTotalCost = Tutal
End Sub

Private Sub txtTotalCost_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtTotalCost_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtVeh_register_CustCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtVINo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtYeer_Validate(Cancel As Boolean)
    If IsNumeric(txtYeer) = False Then: txtYeer = "": Exit Sub
    If IsDate(DateSerial(txtYeer, 1, 1)) = False Then
        txtYeer = ""
    End If

End Sub

