VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCRIS_InquiryVehicleCatalogue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Receiving Entry"
   ClientHeight    =   10005
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   11895
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
   Icon            =   "VehicleCatalogue.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   11895
   Begin VB.PictureBox picBottoms 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   11895
      TabIndex        =   42
      Top             =   9090
      Width           =   11895
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   960
         Top             =   150
      End
      Begin Crystal.CrystalReport rptMRR 
         Left            =   4635
         Top             =   225
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
         Left            =   5895
         ScaleHeight     =   945
         ScaleWidth      =   6075
         TabIndex        =   44
         Top             =   0
         Width           =   6075
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   4890
            MouseIcon       =   "VehicleCatalogue.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "VehicleCatalogue.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   795
            Left            =   4200
            MouseIcon       =   "VehicleCatalogue.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "VehicleCatalogue.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   3510
            MouseIcon       =   "VehicleCatalogue.frx":123A
            MousePointer    =   99  'Custom
            Picture         =   "VehicleCatalogue.frx":138C
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   2820
            MouseIcon       =   "VehicleCatalogue.frx":1686
            MousePointer    =   99  'Custom
            Picture         =   "VehicleCatalogue.frx":17D8
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   2130
            MouseIcon       =   "VehicleCatalogue.frx":1B30
            MousePointer    =   99  'Custom
            Picture         =   "VehicleCatalogue.frx":1C82
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.Label labInventoryStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   780
         Left            =   225
         TabIndex        =   43
         Top             =   90
         Width           =   5325
      End
   End
   Begin VB.PictureBox picTops 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   0
      ScaleHeight     =   1725
      ScaleWidth      =   11895
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.PictureBox picModelDetails 
         BorderStyle     =   0  'None
         Height          =   1800
         Left            =   0
         ScaleHeight     =   1800
         ScaleWidth      =   11955
         TabIndex        =   68
         Top             =   0
         Width           =   11955
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
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   76
            Top             =   900
            Width           =   1680
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
            Left            =   1200
            TabIndex        =   75
            Text            =   "txtDescript"
            Top             =   450
            Width           =   5400
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
            Left            =   4260
            TabIndex        =   74
            Text            =   "Combo1"
            Top             =   1320
            Width           =   2340
         End
         Begin VB.TextBox txtCode 
            BackColor       =   &H80000004&
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
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   60
            Width           =   1185
         End
         Begin VB.TextBox txtMake 
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
            Height          =   345
            Left            =   2445
            TabIndex        =   72
            Top             =   60
            Width           =   4155
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
            Height          =   345
            Left            =   4260
            MaxLength       =   4
            TabIndex        =   71
            Top             =   900
            Width           =   2340
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   900
            Width           =   750
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
            ItemData        =   "VehicleCatalogue.frx":1FE1
            Left            =   1200
            List            =   "VehicleCatalogue.frx":1FEB
            TabIndex        =   69
            Top             =   1320
            Width           =   2415
         End
         Begin MSMask.MaskEdBox txtDateReleased 
            Height          =   345
            Left            =   9645
            TabIndex        =   83
            ToolTipText     =   "Date Vehicles Released"
            Top             =   870
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
            Left            =   9645
            TabIndex        =   84
            Tag             =   "@R"
            ToolTipText     =   "Date Vehicles Received (Recieved Date)"
            Top             =   450
            Width           =   2115
            _ExtentX        =   3731
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
            Format          =   "dd-mmm-yy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtPullOutDate 
            Height          =   345
            Left            =   9645
            TabIndex        =   85
            ToolTipText     =   "Date of Pull Out ( Pull Out Date)"
            Top             =   45
            Width           =   2115
            _ExtentX        =   3731
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
            Format          =   "dd-mmm-yy"
            PromptChar      =   "_"
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
            Left            =   8760
            TabIndex        =   88
            Top             =   495
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
            Left            =   8760
            TabIndex        =   87
            Top             =   915
            Width           =   1185
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
            Left            =   8880
            TabIndex        =   86
            Top             =   90
            Width           =   1185
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
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
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   135
            TabIndex        =   82
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Code / Make"
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
            Left            =   75
            TabIndex        =   81
            Top             =   150
            Width           =   1035
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
            Left            =   600
            TabIndex        =   80
            Top             =   960
            Width           =   510
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
            Left            =   3780
            TabIndex        =   79
            Top             =   960
            Width           =   390
         End
         Begin VB.Label Label26 
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
            Height          =   255
            Left            =   3660
            TabIndex        =   78
            Top             =   1380
            Width           =   705
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
            Left            =   0
            TabIndex        =   77
            Top             =   1380
            Width           =   1170
         End
      End
      Begin VB.PictureBox picRefHeader 
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   8400
         ScaleHeight     =   2115
         ScaleWidth      =   4215
         TabIndex        =   51
         Top             =   75
         Width           =   4215
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   990
         Visible         =   0   'False
         Width           =   1140
      End
   End
   Begin VB.PictureBox picMiddles 
      Align           =   1  'Align Top
      Height          =   7365
      Left            =   0
      ScaleHeight     =   7305
      ScaleWidth      =   11835
      TabIndex        =   16
      Top             =   1725
      Width           =   11895
      Begin VB.VScrollBar ScrollBar1 
         Height          =   5730
         LargeChange     =   500
         Left            =   11520
         SmallChange     =   250
         TabIndex        =   17
         Top             =   0
         Width           =   330
      End
      Begin VB.PictureBox picVehicleReceving 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7365
         Left            =   0
         ScaleHeight     =   7365
         ScaleWidth      =   11580
         TabIndex        =   18
         Top             =   0
         Width           =   11580
         Begin VB.TextBox txtLTOStatus 
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
            Height          =   345
            Left            =   1065
            TabIndex        =   60
            Top             =   4110
            Width           =   6165
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
            Height          =   360
            Left            =   1065
            TabIndex        =   59
            Top             =   2880
            Width           =   6165
         End
         Begin VB.TextBox txtCSR 
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
            Height          =   345
            Left            =   1065
            TabIndex        =   58
            Top             =   4530
            Width           =   1065
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
            Height          =   360
            Left            =   1065
            TabIndex        =   57
            Top             =   3270
            Width           =   6165
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
            Height          =   360
            Left            =   1065
            TabIndex        =   56
            Top             =   3660
            Width           =   6165
         End
         Begin VB.OptionButton optOnShowroom 
            Caption         =   "Units for Display in Showroom"
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
            Left            =   2190
            TabIndex        =   55
            Top             =   4470
            Width           =   3165
         End
         Begin VB.OptionButton optWithProsBuyers 
            Caption         =   "Units with Prospective Buyers"
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
            Left            =   2190
            TabIndex        =   54
            Top             =   4740
            Width           =   3885
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Unknown"
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
            Left            =   2190
            TabIndex        =   53
            Top             =   5340
            Width           =   3465
         End
         Begin VB.OptionButton optReserved 
            Caption         =   "Unit is Reserved"
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
            Left            =   2190
            TabIndex        =   52
            Top             =   5040
            Width           =   3825
         End
         Begin VB.PictureBox picVehicleBees 
            BorderStyle     =   0  'None
            Height          =   2040
            Left            =   120
            ScaleHeight     =   2040
            ScaleWidth      =   11445
            TabIndex        =   39
            Top             =   5700
            Width           =   11445
            Begin MSComctlLib.ListView lstAccesories 
               Height          =   1485
               Left            =   45
               TabIndex        =   40
               Top             =   45
               Width           =   5835
               _ExtentX        =   10292
               _ExtentY        =   2619
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FlatScrollBar   =   -1  'True
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
            Begin MSComctlLib.ListView lstFreeBies 
               Height          =   1485
               Left            =   5895
               TabIndex        =   41
               Top             =   45
               Width           =   5475
               _ExtentX        =   9657
               _ExtentY        =   2619
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FlatScrollBar   =   -1  'True
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
            Height          =   2850
            Left            =   60
            ScaleHeight     =   2850
            ScaleWidth      =   7335
            TabIndex        =   1
            Top             =   0
            Width           =   7335
            Begin VB.TextBox txtFrameNo 
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
               Left            =   5010
               TabIndex        =   10
               Top             =   450
               Width           =   2250
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
               Left            =   1050
               TabIndex        =   3
               Top             =   450
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
               Left            =   1050
               TabIndex        =   2
               Text            =   "Combo1"
               Top             =   75
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
               Left            =   1050
               TabIndex        =   5
               Top             =   1230
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
               Left            =   1050
               TabIndex        =   6
               Top             =   1620
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
               Left            =   1050
               TabIndex        =   7
               Top             =   2040
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
               Left            =   1050
               TabIndex        =   8
               Top             =   2430
               Width           =   2925
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
               Left            =   5010
               TabIndex        =   9
               Top             =   60
               Width           =   2235
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
               Left            =   5010
               TabIndex        =   11
               Top             =   825
               Width           =   2235
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
               Left            =   5010
               TabIndex        =   12
               Top             =   1215
               Width           =   2235
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
               Left            =   5010
               TabIndex        =   13
               Top             =   1605
               Width           =   2235
            End
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
               Left            =   1050
               TabIndex        =   4
               Text            =   "Combo1"
               Top             =   840
               Width           =   2925
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
               Left            =   4010
               TabIndex        =   50
               Top             =   480
               Width           =   810
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Unit"
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
               Left            =   60
               TabIndex        =   21
               Top             =   480
               Width           =   330
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
               Left            =   60
               TabIndex        =   19
               Top             =   90
               Width           =   615
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
               Left            =   60
               TabIndex        =   23
               Top             =   870
               Width           =   450
            End
            Begin VB.Label Label5 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Conduction Sticker"
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
               Height          =   405
               Left            =   60
               TabIndex        =   25
               Top             =   1170
               Width           =   1095
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Prod No"
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
               Left            =   60
               TabIndex        =   27
               Top             =   1650
               Width           =   675
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Serial No"
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
               Left            =   60
               TabIndex        =   28
               Top             =   2070
               Width           =   765
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "VI No"
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
               TabIndex        =   29
               Top             =   2460
               Width           =   435
            End
            Begin VB.Label Label11 
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
               Height          =   255
               Left            =   4010
               TabIndex        =   20
               Top             =   120
               Width           =   1695
            End
            Begin VB.Label Label13 
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
               Height          =   255
               Left            =   4010
               TabIndex        =   22
               Top             =   885
               Width           =   1695
            End
            Begin VB.Label Label14 
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
               Height          =   255
               Left            =   4010
               TabIndex        =   24
               Top             =   1275
               Width           =   1695
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
               Left            =   4010
               TabIndex        =   26
               Top             =   1665
               Width           =   1695
            End
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
            Height          =   5460
            Left            =   7440
            ScaleHeight     =   5460
            ScaleWidth      =   7335
            TabIndex        =   30
            Top             =   120
            Width           =   7335
            Begin VB.TextBox txtNote 
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
               Height          =   1020
               Left            =   0
               TabIndex        =   66
               Top             =   4305
               Width           =   4020
            End
            Begin VB.TextBox txtProfile1 
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
               Height          =   780
               Left            =   0
               TabIndex        =   31
               Top             =   180
               Width           =   4035
            End
            Begin VB.TextBox txtProfile2 
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
               Height          =   780
               Left            =   0
               TabIndex        =   33
               Top             =   1200
               Width           =   4035
            End
            Begin VB.TextBox txtProfile3 
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
               Height          =   780
               Left            =   0
               TabIndex        =   35
               Top             =   2220
               Width           =   4035
            End
            Begin VB.TextBox txtProfile4 
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
               Height          =   780
               Left            =   0
               TabIndex        =   37
               Top             =   3240
               Width           =   4035
            End
            Begin VB.Label Label40 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Inventory Note:"
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
               Left            =   30
               TabIndex        =   67
               Top             =   4080
               Width           =   1665
            End
            Begin VB.Label Label34 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Profile 1"
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
               Left            =   0
               TabIndex        =   32
               Top             =   -30
               Width           =   1065
            End
            Begin VB.Label Label35 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Profile 2"
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
               Left            =   0
               TabIndex        =   34
               Top             =   960
               Width           =   1065
            End
            Begin VB.Label Label36 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Profile 3"
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
               Left            =   0
               TabIndex        =   36
               Top             =   1980
               Width           =   1065
            End
            Begin VB.Label Label37 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Profile 4"
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
               Left            =   0
               TabIndex        =   38
               Top             =   3000
               Width           =   1065
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
            Left            =   120
            TabIndex        =   65
            Top             =   4170
            Width           =   945
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks 1"
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
            TabIndex        =   64
            Top             =   2910
            Width           =   930
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
            Left            =   120
            TabIndex        =   63
            Top             =   4560
            Width           =   360
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks 2"
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
            TabIndex        =   62
            Top             =   3300
            Width           =   930
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks 3"
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
            TabIndex        =   61
            Top             =   3690
            Width           =   930
         End
      End
   End
End
Attribute VB_Name = "frmCRIS_InquiryVehicleCatalogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================================================
'FUNCTION / FEATURE :ADDED MODEL CODE THINGS FOR
'DATE STARTED       :6/5/200713:06
'LAST UPDATED       :6/5/200713:06
'DATABASE UPDATES   :LOOK UP IN ALL_MODELCODE
'WHO UPDATED        :AXP  6/5/2007
'UDPATING CODE    : AXP-652005106
'==========================================================================================
'FUNCTION / FEATURE :   ADDED TRASMISSION TYPE DEFAULT VALUE PARSING BY DESCRIPTION HOWEVER USER CAN SELECT IT TOO
'                   :   IN Order to Compelete Vehcile Check List Form Easy Transmission Is Added
'DATE STARTED       :   6/7/200717:43
'LAST UPDATED       :   6/7/200717:43
'DATABASE UPDATES   :
'WHO UPDATED        :   AXP 672007543
'UDPATING CODE      :   AXP-672007543
'==========================================================================================
'FUNCTION / FEATURE :   ADDED PO FOR MRR DETAILS PROCESS FROM PO IT WILL BE SAVED TO MRR
' AUTOMATION OF PO TO MRR
'DATE STARTED       :   6/13/200713:49
'LAST UPDATED       :   6/13/200713:49
'DATABASE UPDATES   :
'WHO UPDATED        :   AXP 06132007149
'UDPATING CODE      :   AXP-06132007149
'==========================================================================================

Option Explicit
Dim rsMRRINV                           As ADODB.Recordset
















Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub cmdFind_Click()

frmCRIS_SearchVehicleInfo.Show
End Sub

Private Sub cmdNext_Click()
    rsMRRINV.MoveNext
    If rsMRRINV.EOF Then
        rsMRRINV.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub


Private Sub cmdPrevious_Click()
    rsMRRINV.MovePrevious
    If rsMRRINV.BOF Then
        rsMRRINV.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    PrintSQLReport rptMRR, SMIS_REPORT_PATH & "mrr.rpt", "{MRRINV.ID} = " & labid, DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub




Private Sub Form_Load()
    Screen.MousePointer = 11
    Me.Height = Screen.TwipsPerPixelY * 550
    picMiddles.Height = Me.ScaleHeight - picTops.Height - picBottoms.Height
    ScrollBar1.Height = picMiddles.ScaleHeight - 15
    ScrollBar1.Max = Abs(picMiddles.ScaleHeight - picVehicleReceving.Height) + 20
    CenterMe frmMain, Me, 1

SEARCH_TAB = 0
    rsRefresh
    InitData
    InitMemVars

    picTops.Enabled = False
    picVehicleReceving.Enabled = False
    
    picAdds.Visible = True
    StoreMemvars
    Screen.MousePointer = 0

End Sub
Sub InitData()
    Call AddColumnHeader("DESCRIPTION", lstAccesories)
    Call AddColumnHeader("DESCRIPTION", lstFreeBies)
    Call ResizeColumnHeader(lstAccesories, "90")
    Call ResizeColumnHeader(lstFreeBies, "90")
 

End Sub
Sub InitMemVars()
    labInventoryStatus = ""
    txtCode.Text = ""
    cboModelDescript.Text = ""
    txtMake.Text = "Hyundai"
    txtModel.Text = ""
    txtYeer.Text = ""
    
    
    cboSource.Clear
    cboSource.AddItem "HARI"
    cboSource.Text = "HARI"

    txtUnit.Text = ""



    txtIgnKey.Text = ""
    txtProdNo.Text = ""
    txtSerialNo.Text = ""
    txtVINo.Text = ""
    txtEngineNo.Text = ""
    txtFuelUsed.Text = ""
    txtPistonDisp.Text = ""
    txtGVW.Text = ""
    
    
    
    txtDateReceived.Text = FormatDateTime(Now, vbShortDate)
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
    lstAccesories.ListItems.Clear
    lstFreeBies.ListItems.Clear
    cboColor = ""
    txtFrameNo = ""
    optOnShowroom.Value = True
    optWithProsBuyers.Value = False
End Sub



Private Sub rsRefresh()
    Set rsMRRINV = New ADODB.Recordset
    rsMRRINV.CursorLocation = adUseClient
    Call rsMRRINV.Open("SELECT * from SMIS_MrrInv order by Datereceived DESC", gconDMIS, adOpenKeyset)
End Sub



Private Sub ScrollBar1_Change()
    picVehicleReceving.Top = 0 - ScrollBar1.Value
End Sub


Private Sub StoreMemvars()
    If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then

        labid.Caption = rsMRRINV!id
        txtCode = Null2String(rsMRRINV!CODE)
        
        cboModelDescript = Null2String(rsMRRINV!DESCRIPT)
        txtMake = Null2String(rsMRRINV!Make)
        txtModel = Null2String(rsMRRINV!Model)
        txtModelCode = Null2String(rsMRRINV!ModelCode)
        SetClass
        'cboClass.ListIndex = SelectCombo(cboClass, Null2String(rsMRRINV!Class))
        txtYeer = Null2String(rsMRRINV!YEER)
        cboSource = Null2String(rsMRRINV!Source)
        txtUnit = Null2String(rsMRRINV!unit)
        cboColor = Null2String(rsMRRINV!Color)
        txtIgnKey = Null2String(rsMRRINV!IGNKEY)
        txtProdNo = Null2String(rsMRRINV!ProdNo)
        cboTransmission = Null2String(rsMRRINV!Transmission)


        txtSerialNo = Null2String(rsMRRINV!SERIALNO)
        txtVINo = Null2String(rsMRRINV!VINO)
        txtEngineNo = Null2String(rsMRRINV!ENGINENO)
        txtFuelUsed = Null2String(rsMRRINV!fuelused)
        txtPistonDisp = Null2String(rsMRRINV!pistondisp)
        txtFrameNo = Null2String(rsMRRINV!FrameNo)
        txtGVW = Null2String(rsMRRINV!gvw)

        txtPullOutDate = Null2String(rsMRRINV!PullOutDate)

        txtDateReceived = Null2String(rsMRRINV!DateReceived)
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
        Dim rsInvStatus                As ADODB.Recordset
        Dim Temprs  As ADODB.Recordset
        RELEASEINFO = Null2String(rsMRRINV!DateReleased)
        ISTATUS = Null2String(rsMRRINV!ISTATUS)
        'sold and unrealease
        If ISTATUS = "S" And IsDate(RELEASEINFO) = False Then
            Set rsInvStatus = gconDMIS.Execute("Select CUSNAM  from ALL_CUSMAS WHERE CUSCDE=" & N2Str2Null(rsMRRINV!CustomerCode))
            If Not (rsInvStatus.EOF Or rsInvStatus.BOF) Then
                labInventoryStatus = "** INVOICED / NOT RELEASED **" & Null2String(rsInvStatus!CUSNAM)
            Else
                labInventoryStatus = "** INVOICED / NOT RELEASED ** CUSTOMER INFORMATION MISSING"
            End If
            picVehicleDetails.Enabled = False
            picModelDetails.Enabled = False
            picRefHeader.Enabled = False
            'sold and released
        ElseIf ISTATUS = "R" And IsDate(RELEASEINFO) = True Then
            Set rsInvStatus = gconDMIS.Execute("Select CUSNAM  from ALL_CUSMAS WHERE CUSCDE=" & N2Str2Null(rsMRRINV!CustomerCode))
            If Not (rsInvStatus.EOF Or rsInvStatus.BOF) Then
                labInventoryStatus = "** SOLD  TO **" & Null2String(rsInvStatus!CUSNAM)
            Else
                labInventoryStatus = "** SOLD BUT CUSTOMER INFORMATION MISSING"
            End If
            picVehicleDetails.Enabled = False
            picModelDetails.Enabled = False
            picRefHeader.Enabled = False
            'allocated
        ElseIf ISTATUS = "A" Then
            Set rsInvStatus = gconDMIS.Execute("Select CUSNAM  from ALL_CUSMAS WHERE CUSCDE=" & N2Str2Null(rsMRRINV!CustomerCode))
            If Not (rsInvStatus.EOF Or rsInvStatus.BOF) Then
                labInventoryStatus = "** ALLOCATED FOR **" & Null2String(rsInvStatus!CUSNAM)
            Else
                labInventoryStatus = "**** ALLOCATED / CUSTOMER INFORMATION MISSING**"
            End If
            picVehicleDetails.Enabled = False
            picModelDetails.Enabled = False
            picRefHeader.Enabled = False
        ElseIf ISTATUS = "D" Then
            labInventoryStatus = "**DEMO VEHICLE**"

        ElseIf ISTATUS = "T" Then
            
            Set Temprs = gconDMIS.Execute("Select Entity_From From SMIS_StockTransfer where VSNO=" & N2Str2Null(rsMRRINV!ProdNo))
            If Not (Temprs.EOF Or Temprs.BOF) Then
            labInventoryStatus = "**UNIT TRANSFERED FROM " & Null2String(Temprs!Entity_From) & " **"
            Else
            labInventoryStatus = "**UNIT TRANSFERED MISSING TRANSFEREE INFO**"
            End If
            
            
             picVehicleDetails.Enabled = False
            picModelDetails.Enabled = False
            picRefHeader.Enabled = False

        Else
            picRefHeader.Enabled = True
            picVehicleDetails.Enabled = True
            picModelDetails.Enabled = True
            labInventoryStatus = "** AVAILABLE/OPEN**"

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


        flex_FillListView gconDMIS.Execute("Select Description, COST from SMIS_MRRINV_DETAIL WHERE IsFree=0 AND IgnKeyNo=" & N2Str2Null(txtIgnKey)), lstAccesories
        flex_FillListView gconDMIS.Execute("Select Description, COST  from SMIS_MRRINV_DETAIL WHERE IsFree=1 AND IgnKeyNo=" & N2Str2Null(txtIgnKey)), lstFreeBies


    Else
        ShowNoRecord

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

End Sub










Sub SearchID(xxx)

    Dim varBookMark                    As Variant
    varBookMark = rsMRRINV.Bookmark
    rsMRRINV.MoveFirst
    rsMRRINV.Find "id = " & xxx
    If (rsMRRINV.BOF = True) Or (rsMRRINV.EOF = True) Then
        MsgBox "Record not found"
        rsMRRINV.Bookmark = varBookMark
    End If

    StoreMemvars
End Sub

Function GetClassCode() As String
    Dim Temprs                         As ADODB.Recordset

    If cboClass.ListIndex <> -1 Then

        Set Temprs = gconDMIS.Execute("SELECT CODE FROM SMIS_VehiclesClass Where ID= " & cboClass.ItemData(cboClass.ListIndex))

        If Not (Temprs.EOF Or Temprs.BOF) Then
            GetClassCode = Null2String(Temprs!CODE)
        End If

        Set Temprs = Nothing

    Else
        GetClassCode = vbNullString
    End If
End Function




Sub SetClass()
    Dim Temprs                         As ADODB.Recordset


    Set Temprs = gconDMIS.Execute("SELECT ClassName FROM SMIS_VehiclesClass Where Code= " & N2Str2Null(rsMRRINV!Class))

    If Not (Temprs.EOF Or Temprs.BOF) Then
        cboClass.ListIndex = SelectCombo(cboClass, Null2String(Temprs!ClassName))
    End If

    Set Temprs = Nothing


End Sub

Function DetectATMT(strx)
    Dim I                              As Integer
    Dim ax
    ax = Split(strx)
    For I = 1 To UBound(ax)
        If InStr(1, ax(I), "MT") > 0 Then
            DetectATMT = "MT"
            Exit Function
        End If
    Next
    DetectATMT = "AT"
    Erase ax
End Function

