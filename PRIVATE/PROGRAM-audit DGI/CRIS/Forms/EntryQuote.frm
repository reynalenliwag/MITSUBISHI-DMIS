VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmCRIS_EntryQuotation 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9090
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
   Icon            =   "EntryQuote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picServiceDetails 
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
      Height          =   3840
      Left            =   1410
      ScaleHeight     =   3810
      ScaleWidth      =   8295
      TabIndex        =   59
      Top             =   2490
      Visible         =   0   'False
      Width           =   8325
      Begin VB.CommandButton cmdCancelDetailService 
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
         Left            =   7980
         TabIndex        =   111
         Top             =   0
         Width           =   285
      End
      Begin VB.PictureBox picServiceDetailView 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   3465
         Left            =   30
         ScaleHeight     =   3435
         ScaleWidth      =   3990
         TabIndex        =   95
         Top             =   300
         Width           =   4020
         Begin VB.Label zlblC 
            Appearance      =   0  'Flat
            BackColor       =   &H00750A04&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "::: SERVICE DETAILS :::"
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
            Index           =   16
            Left            =   0
            TabIndex        =   108
            Top             =   0
            Width           =   5205
         End
         Begin VB.Label lblJOpCode 
            BackColor       =   &H00E0E0E0&
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
            Height          =   270
            Left            =   1545
            TabIndex        =   107
            Top             =   2310
            Width           =   2415
         End
         Begin VB.Label lblJModel 
            BackColor       =   &H00E0E0E0&
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
            Height          =   270
            Left            =   1545
            TabIndex        =   106
            Top             =   2010
            Width           =   2415
         End
         Begin VB.Label lblJDescript 
            BackColor       =   &H00E0E0E0&
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
            Height          =   1140
            Left            =   30
            TabIndex        =   105
            Top             =   840
            Width           =   3900
         End
         Begin VB.Label lblJcode 
            BackColor       =   &H00E0E0E0&
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
            Height          =   270
            Left            =   1545
            TabIndex        =   104
            Top             =   300
            Width           =   2415
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "OP Code"
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
            Left            =   45
            TabIndex        =   103
            Top             =   2310
            Width           =   1470
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "Job Model:"
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
            Left            =   45
            TabIndex        =   102
            Top             =   2010
            Width           =   1470
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00ECCABD&
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
            Index           =   17
            Left            =   45
            TabIndex        =   101
            Top             =   600
            Width           =   3900
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "Job Code"
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
            Index           =   16
            Left            =   45
            TabIndex        =   100
            Top             =   300
            Width           =   1470
         End
         Begin VB.Label lblJFlatRate 
            BackColor       =   &H00E0E0E0&
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
            Height          =   270
            Left            =   1545
            TabIndex        =   99
            Top             =   2610
            Width           =   2415
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "Flat Rate"
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
            Index           =   15
            Left            =   45
            TabIndex        =   98
            Top             =   2610
            Width           =   1470
         End
         Begin VB.Label lblJStd_MHRS 
            BackColor       =   &H00E0E0E0&
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
            Height          =   480
            Left            =   1545
            TabIndex        =   97
            Top             =   2910
            Width           =   2415
         End
         Begin VB.Label lblCapDetails 
            BackColor       =   &H00ECCABD&
            Caption         =   "Standard Man Hours"
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
            Height          =   495
            Index           =   14
            Left            =   45
            TabIndex        =   96
            Top             =   2910
            Width           =   1470
         End
      End
      Begin VB.ComboBox cboJobCodes 
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
         Left            =   4080
         TabIndex        =   94
         Top             =   1170
         Width           =   1275
      End
      Begin VB.ComboBox cboJobCategory 
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
         Left            =   4065
         TabIndex        =   71
         Text            =   "cboProductList"
         Top             =   540
         Width           =   4110
      End
      Begin VB.CommandButton cmdOkServices 
         Caption         =   "OK"
         Height          =   375
         Left            =   6780
         TabIndex        =   62
         Top             =   3270
         Width           =   645
      End
      Begin VB.CommandButton cmdCancelDetailService 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   0
         Left            =   7500
         TabIndex        =   61
         Top             =   3270
         Width           =   645
      End
      Begin VB.ComboBox cboServiceList 
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
         Left            =   5370
         TabIndex        =   60
         Top             =   1170
         Width           =   2805
      End
      Begin EditLib.fpCurrency txtSAmount 
         Height          =   345
         Left            =   4065
         TabIndex        =   63
         ToolTipText     =   "Total Amount of Bill Line"
         Top             =   2385
         Width           =   4095
         _Version        =   196608
         _ExtentX        =   7223
         _ExtentY        =   609
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   7670276
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   0
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   0
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   2
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   192
         InvalidOption   =   0
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   2
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9999999"
         MinValue        =   "0"
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpLongInteger txtSQty 
         Height          =   345
         Left            =   4065
         TabIndex        =   64
         Top             =   1800
         Width           =   1185
         _Version        =   196608
         _ExtentX        =   2090
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   7670276
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   1
         ButtonWidth     =   0
         ButtonWrap      =   0   'False
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   1
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   1
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "0"
         MaxValue        =   "9999"
         MinValue        =   "0"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency txtSRate 
         Height          =   345
         Left            =   5415
         TabIndex        =   65
         ToolTipText     =   "Rate of Material"
         Top             =   1800
         Width           =   2715
         _Version        =   196608
         _ExtentX        =   4789
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   7670276
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   0
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   0
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   2
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   192
         InvalidOption   =   1
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   2
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9999999"
         MinValue        =   "0"
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
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
         Left            =   4065
         TabIndex        =   72
         Top             =   315
         Width           =   795
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Name:"
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
         Left            =   4065
         TabIndex        =   70
         Top             =   900
         Width           =   1170
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Index           =   14
         Left            =   4065
         TabIndex        =   69
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
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
         Left            =   4065
         TabIndex        =   68
         Top             =   1530
         Width           =   270
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Index           =   11
         Left            =   5415
         TabIndex        =   67
         Top             =   1575
         Width           =   360
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Left            =   -15
         TabIndex        =   66
         Top             =   0
         Width           =   8295
         _Version        =   655364
         _ExtentX        =   14631
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "::: Add Service :::"
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
      Left            =   1950
      ScaleHeight     =   5565
      ScaleWidth      =   6990
      TabIndex        =   33
      Top             =   1620
      Visible         =   0   'False
      Width           =   7020
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
         TabIndex        =   110
         Top             =   30
         Width           =   285
      End
      Begin VB.PictureBox picVehiclesDetail 
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
         TabIndex        =   74
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
            TabIndex        =   127
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
            TabIndex        =   126
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
            TabIndex        =   93
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
            TabIndex        =   92
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
            TabIndex        =   91
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
            TabIndex        =   90
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
            TabIndex        =   89
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
            TabIndex        =   88
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
            TabIndex        =   87
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
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   84
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
            TabIndex        =   83
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
            TabIndex        =   82
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
            TabIndex        =   81
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
            TabIndex        =   80
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
            TabIndex        =   79
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
            TabIndex        =   78
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
            TabIndex        =   77
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
            TabIndex        =   76
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
            TabIndex        =   75
            Top             =   -30
            Width           =   3555
         End
      End
      Begin EditLib.fpDoubleSingle txtVAOR 
         Height          =   345
         Left            =   3645
         TabIndex        =   56
         Top             =   3510
         Width           =   3210
         _Version        =   196608
         _ExtentX        =   5662
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   -2147483635
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "0"
         DecimalPlaces   =   -1
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         Picture         =   "EntryQuote.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   35
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
         Picture         =   "EntryQuote.frx":06F6
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5085
         Width           =   420
      End
      Begin EditLib.fpCurrency txtVBalToFin 
         Height          =   345
         Left            =   3645
         TabIndex        =   49
         ToolTipText     =   "Total Amount of Bill Line"
         Top             =   4050
         Width           =   3210
         _Version        =   196608
         _ExtentX        =   5662
         _ExtentY        =   609
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   7670276
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   0
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   0
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   2
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   192
         InvalidOption   =   0
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   2
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9999999"
         MinValue        =   "0"
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpLongInteger txtVQty 
         Height          =   345
         Left            =   6345
         TabIndex        =   50
         Top             =   585
         Width           =   555
         _Version        =   196608
         _ExtentX        =   979
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   7670276
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   1
         ButtonWidth     =   0
         ButtonWrap      =   0   'False
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   1
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   1
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "0"
         MaxValue        =   "9999"
         MinValue        =   "1"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency txtVdownpayment 
         Height          =   345
         Left            =   3645
         TabIndex        =   51
         ToolTipText     =   "Rate of Material"
         Top             =   2340
         Width           =   3210
         _Version        =   196608
         _ExtentX        =   5662
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   7670276
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   0
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   0
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   2
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   192
         InvalidOption   =   1
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   2
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "1000000000000000000000"
         MinValue        =   "0"
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency txtvNetMonthlyMort 
         Height          =   345
         Left            =   3645
         TabIndex        =   52
         ToolTipText     =   "Total Amount of Bill Line"
         Top             =   4680
         Width           =   3210
         _Version        =   196608
         _ExtentX        =   5662
         _ExtentY        =   609
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   7670276
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   0
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   0
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   2
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   192
         InvalidOption   =   0
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   2
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9999999"
         MinValue        =   "0"
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency txtVNetRate 
         Height          =   345
         Left            =   3645
         TabIndex        =   53
         ToolTipText     =   "Rate of Material"
         Top             =   1170
         Width           =   3210
         _Version        =   196608
         _ExtentX        =   5662
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   7670276
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   2
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   192
         InvalidOption   =   1
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   2
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "1000000000000000000000"
         MinValue        =   "0"
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   3
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency txtVTotalAmount 
         Height          =   345
         Left            =   3645
         TabIndex        =   54
         ToolTipText     =   "Rate of Material"
         Top             =   1755
         Width           =   3210
         _Version        =   196608
         _ExtentX        =   5662
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   7670276
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   0
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   0
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   2
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   192
         InvalidOption   =   1
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   2
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "1000000000000000000000"
         MinValue        =   "0"
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin XtremeShortcutBar.ShortcutCaption caption 
         Height          =   285
         Left            =   0
         TabIndex        =   55
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
         TabIndex        =   57
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   2115
         Width           =   1275
      End
   End
   Begin VB.PictureBox picProductDetails 
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
      Height          =   3105
      Left            =   1860
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   3075
      ScaleWidth      =   5370
      TabIndex        =   23
      Top             =   2190
      Visible         =   0   'False
      Width           =   5400
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
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   109
         Top             =   30
         Width           =   285
      End
      Begin VB.TextBox txtProductCode 
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
         Height          =   285
         Left            =   270
         TabIndex        =   73
         Top             =   810
         Width           =   1410
      End
      Begin VB.ComboBox cboProductList 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   810
         Width           =   2895
      End
      Begin VB.CommandButton cmdCancelDetailProduct 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   0
         Left            =   4005
         TabIndex        =   32
         Top             =   2520
         Width           =   645
      End
      Begin VB.CommandButton cmdOkMaterials 
         Caption         =   "OK"
         Height          =   375
         Left            =   3285
         TabIndex        =   31
         Top             =   2520
         Width           =   645
      End
      Begin EditLib.fpCurrency txtAmount 
         Height          =   345
         Left            =   2250
         TabIndex        =   24
         Tag             =   "@ZERO"
         ToolTipText     =   "Total Amount of Bill Line"
         Top             =   2070
         Width           =   2385
         _Version        =   196608
         _ExtentX        =   4207
         _ExtentY        =   609
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   7670276
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   0
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   0
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   2
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   192
         InvalidOption   =   0
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   2
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9999999"
         MinValue        =   "0"
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpLongInteger txtQty 
         Height          =   345
         Left            =   3375
         TabIndex        =   25
         Tag             =   "@ZERO"
         Top             =   1215
         Width           =   1185
         _Version        =   196608
         _ExtentX        =   2090
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   7670276
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   1
         ButtonWidth     =   0
         ButtonWrap      =   0   'False
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   1
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   1
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "0"
         MaxValue        =   "9999"
         MinValue        =   "0"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency txtRate 
         Height          =   345
         Left            =   2250
         TabIndex        =   26
         Tag             =   "@ZERO"
         ToolTipText     =   "Rate of Material"
         Top             =   1665
         Width           =   2355
         _Version        =   196608
         _ExtentX        =   4154
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   7670276
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   0
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   0
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   2
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   192
         InvalidOption   =   1
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   2
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9999999"
         MinValue        =   "0"
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin XtremeShortcutBar.ShortcutCaption cap2 
         Height          =   330
         Left            =   0
         TabIndex        =   58
         Top             =   0
         Width           =   5415
         _Version        =   655364
         _ExtentX        =   9551
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "::: Add Product :::"
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
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Index           =   3
         Left            =   225
         TabIndex        =   30
         Top             =   1755
         Width           =   360
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
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
         Left            =   270
         TabIndex        =   29
         Top             =   1305
         Width           =   270
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   225
         TabIndex        =   28
         Top             =   2115
         Width           =   660
      End
      Begin VB.Label lblQuotationParticular 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name:"
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
         Left            =   180
         TabIndex        =   27
         Top             =   480
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000A&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9420
      Picture         =   "EntryQuote.frx":0900
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8400
      Width           =   660
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H8000000A&
      DisabledPicture =   "EntryQuote.frx":1874
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8730
      MaskColor       =   &H00FFFFFF&
      Picture         =   "EntryQuote.frx":27E8
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8400
      Width           =   660
   End
   Begin VB.PictureBox picTop 
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
      ScaleWidth      =   10200
      TabIndex        =   1
      Top             =   330
      Width           =   10200
      Begin VB.CommandButton cmdAddMaterials 
         Caption         =   "Add Materials"
         Height          =   330
         Left            =   4095
         TabIndex        =   22
         Top             =   900
         Width           =   1230
      End
      Begin VB.CommandButton cmdAddService 
         Caption         =   "Add Services"
         Height          =   330
         Left            =   4095
         TabIndex        =   21
         Top             =   2025
         Width           =   1230
      End
      Begin VB.CommandButton cmdAddParts 
         Caption         =   "Add Parts"
         Height          =   330
         Left            =   4095
         TabIndex        =   20
         Top             =   1650
         Width           =   1230
      End
      Begin VB.CommandButton cmdAddVehicles 
         Caption         =   "Add Vehicles"
         Height          =   330
         Left            =   4095
         TabIndex        =   19
         Top             =   1275
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
         Left            =   135
         MaxLength       =   220
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   900
         Width           =   3720
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
            TabIndex        =   134
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
            Index           =   0
            Left            =   1500
            TabIndex        =   13
            Top             =   2085
            Width           =   2805
         End
         Begin VB.Label lblContactNo 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Index           =   0
            Left            =   1500
            TabIndex        =   12
            Top             =   1800
            Width           =   2805
         End
         Begin VB.Label lblAddress 
            BackColor       =   &H00C0C0C0&
            Height          =   690
            Index           =   0
            Left            =   15
            TabIndex        =   11
            Top             =   1095
            Width           =   4290
         End
         Begin VB.Label lblAccountName 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Index           =   0
            Left            =   1500
            TabIndex        =   10
            Top             =   570
            Width           =   2805
         End
         Begin VB.Label lblCustomerName 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Index           =   0
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
            ForeColor       =   &H8000000D&
            Height          =   270
            Index           =   0
            Left            =   15
            TabIndex        =   4
            Top             =   285
            Width           =   1470
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4275
         Top             =   90
      End
      Begin EditLib.fpText txtQuotationCode 
         Height          =   345
         Left            =   150
         TabIndex        =   0
         Tag             =   "0"
         Top             =   315
         Width           =   3705
         _Version        =   196608
         _ExtentX        =   6535
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   0
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   -1  'True
         AutoCase        =   1
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   "~^!`<>/?"";+="
         MaxLength       =   6
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   2
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   0
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label zlblC 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00750A04&
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   675
         Width           =   945
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
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8565
      Left            =   0
      ScaleHeight     =   8535
      ScaleWidth      =   2625
      TabIndex        =   128
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
      Begin XtremeReportControl.ReportControl lvSearch 
         Height          =   6765
         Left            =   90
         TabIndex        =   135
         Top             =   690
         Width           =   2475
         _Version        =   655364
         _ExtentX        =   4366
         _ExtentY        =   11933
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.TextBox txtSearchQuotation 
         Height          =   345
         Left            =   90
         TabIndex        =   129
         Top             =   300
         Width           =   2445
      End
      Begin EditLib.fpDoubleSingle txtTotalAmount 
         Height          =   420
         Left            =   60
         TabIndex        =   131
         Top             =   7800
         Width           =   2520
         _Version        =   196608
         _ExtentX        =   4445
         _ExtentY        =   741
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   8388608
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   0
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   0
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   0
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   0
         AlignTextH      =   2
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   1
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   3
         MarginBottom    =   0
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   1
         Text            =   "0.00"
         DecimalPlaces   =   2
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "99999999999"
         MinValue        =   "0"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   0.25
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label zlblC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter"
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
         Index           =   17
         Left            =   90
         TabIndex        =   132
         Top             =   0
         Width           =   420
      End
      Begin VB.Label zlblC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00750A04&
         Height          =   255
         Index           =   12
         Left            =   150
         TabIndex        =   130
         Top             =   7590
         Width           =   1545
      End
   End
   Begin VB.PictureBox picDetails 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5655
      ScaleWidth      =   10245
      TabIndex        =   14
      Top             =   2820
      Width           =   10245
      Begin XtremeReportControl.ReportControl lvQuotationVehicles 
         Height          =   2355
         Left            =   3390
         TabIndex        =   16
         Top             =   3090
         Width           =   6615
         _Version        =   655364
         _ExtentX        =   11668
         _ExtentY        =   4154
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         AllowColumnResize=   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeReportControl.ReportControl lvQuotation 
         Height          =   3075
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   10005
         _Version        =   655364
         _ExtentX        =   17648
         _ExtentY        =   5424
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnRemove=   0   'False
         AllowColumnReorder=   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picVehiclesQDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   2355
         Left            =   0
         ScaleHeight     =   2325
         ScaleWidth      =   3330
         TabIndex        =   112
         Top             =   3090
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
            TabIndex        =   125
            Top             =   1710
            Width           =   1470
         End
         Begin VB.Label lblVQNetMonthly 
            BackColor       =   &H00C0C0C0&
            Height          =   600
            Left            =   1485
            TabIndex        =   124
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
            TabIndex        =   123
            Top             =   855
            Width           =   1470
         End
         Begin VB.Label lblVQTerms 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1485
            TabIndex        =   122
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
            TabIndex        =   121
            Top             =   576
            Width           =   1470
         End
         Begin VB.Label lblVQDownPayment 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1485
            TabIndex        =   120
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
            TabIndex        =   119
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
            TabIndex        =   118
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
            TabIndex        =   117
            Top             =   1425
            Width           =   1470
         End
         Begin VB.Label lblVQCode 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1485
            TabIndex        =   116
            Top             =   285
            Width           =   2805
         End
         Begin VB.Label lblVQAOR 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1485
            TabIndex        =   115
            Top             =   1140
            Width           =   2805
         End
         Begin VB.Label lblVQBalToFinanced 
            BackColor       =   &H00C0C0C0&
            Height          =   270
            Left            =   1485
            TabIndex        =   114
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
            TabIndex        =   113
            Top             =   0
            Width           =   5205
         End
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption capMain 
      Height          =   315
      Left            =   0
      TabIndex        =   133
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
Attribute VB_Name = "frmCRIS_EntryQuotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QuotationID                              As Long
Dim QuotationType                            As String
Dim ProfileID                                As Long
Dim ProfileType                              As String
Dim ProspectID                               As Long
Dim AcctName                                 As String

Dim ListEdit                                 As Boolean

Dim RsParts_N_Mat                            As Recordset
Dim RsJobs                                   As Recordset
Dim RsVehicles                               As Recordset



Event ChangeList(Changed As Boolean)

Private Sub cboJobCategory_Click()
    Dim ID                                   As Long
    Dim Particular                           As String
    Dim code                                 As String
    RsJobs.Filter = "Category='" & cboJobCategory.Text & "'"
    cboServiceList.Clear
    While Not RsJobs.EOF
        ID = RsJobs!ID
        Particular = RsJobs!Particulars
        code = RsJobs!code
        cboJobCodes.AddItem (code)
        cboJobCodes.ItemData(cboJobCodes.NewIndex) = ID
        cboServiceList.AddItem (Particular)
        cboServiceList.ItemData(cboServiceList.NewIndex) = ID
        RsJobs.MoveNext
    Wend
End Sub

Private Sub cboJobCodes_LostFocus()
    cboServiceList.ListIndex = SelectCombo(cboJobCodes, cboJobCodes.Text, False)
    LabelJobDescriptions
End Sub

Private Sub cboProductList_Click()
    If ListEdit = False Then
        RsParts_N_Mat.Filter = "ID=" & cboProductList.ItemData(cboProductList.ListIndex)
        If Not (RsParts_N_Mat.EOF Or RsParts_N_Mat.BOF) Then
            txtProductCode.Text = RsParts_N_Mat.Fields("CODE").Value
            txtQty.Value = 1
            txtRate.Value = RsParts_N_Mat.Fields("Rate").Value
        End If
    End If
End Sub

Private Sub cboServiceList_Click()
    RsJobs.Filter = "ID=" & cboServiceList.ItemData(cboServiceList.ListIndex)
    If Not (RsJobs.EOF Or RsJobs.BOF) Then
        txtProductCode.Text = RsJobs.Fields("CODE").Value
        txtSQty.Value = 1
        txtSRate.Value = RsJobs.Fields("Rate").Value
        txtSAmount.Value = txtSRate.Value * txtSQty.Value
    End If
End Sub

Private Sub cboVehicles_Click()
    Dim temprs                               As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT * From CRIS_MRRINV WHERE ID=" & cboVehicles.ItemData(cboVehicles.ListIndex))
    If Not (temprs.EOF Or temprs.BOF) Then
        txtVNetRate.Value = N2Str2IntZero(temprs!taggedprice)
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
    Dim NoOfMonths                           As Integer
    Dim AOR                                  As Double
    AOR = cboVTerm.ItemData(cboVTerm.ListIndex)
    NoOfMonths = Left(cboVTerm.Text, 2)
    If NoOfMonths = 0 Then AOR = 0: NoOfMonths = 1
    If NoOfMonths = 12 Then AOR = 7.61
    If NoOfMonths = 18 Then AOR = 10.48
    If NoOfMonths = 24 Then AOR = 17.45
    If NoOfMonths = 36 Then AOR = 25.55
    If NoOfMonths = 48 Then AOR = 33.96
    If NoOfMonths = 60 Then AOR = 44.15
    txtVAOR.Value = AOR
    txtVBalToFin.Value = txtVTotalAmount.Value - txtVdownpayment.Value
    txtvNetMonthlyMort.Value = ((txtVBalToFin.Value) * (1 + (AOR / 100)) / NoOfMonths)
End Sub

Sub CenterPicture(picx As PictureBox)
    picx.Left = (Me.ScaleWidth - picx.Width) / 2
    picx.Top = (Me.ScaleHeight - picx.Height) / 2
End Sub

Private Sub cmdAddMaterials_Click()
    ListEdit = False
    RsParts_N_Mat.Filter = "TYPE='M'"
    QuotationType = "Materials"
    cboProductList.Clear
    cboProductList.Enabled = True
    txtProductCode.Enabled = True
    cap2.caption = "Add Materials"
    While Not RsParts_N_Mat.EOF
        cboProductList.AddItem (RsParts_N_Mat.Fields("Particulars"))
        cboProductList.ItemData(cboProductList.NewIndex) = RsParts_N_Mat.Collect(0)
        RsParts_N_Mat.MoveNext
    Wend
    cmdSave.Enabled = True
    ShowForm picProductDetails.hwnd
End Sub

Private Sub cmdAddParts_Click()
    ListEdit = False
    QuotationType = "Parts"
    cboProductList.Enabled = True
    txtProductCode.Enabled = True
    RsParts_N_Mat.Filter = "TYPE='P'"
    cap2.caption = "Add Parts"
    FillListData cboProductList, RsParts_N_Mat, "Particulars"
    ShowForm picProductDetails.hwnd
End Sub

Private Sub cmdAddService_Click()
    ListEdit = False
    QuotationType = "Services"
    ShowForm picServiceDetails.hwnd
End Sub

Private Sub cmdAddVehicles_Click()
    ListEdit = False
    QuotationType = "Vehicles"
    ShowForm picVehiclesDetails.hwnd
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdCancelDetailProduct_Click(Index As Integer)
    MyPictureState
End Sub

Private Sub cmdCancelDetailService_Click(Index As Integer)
    MyPictureState
End Sub

Private Sub cmdCancelDetailVehicles_Click(Index As Integer)
    MyPictureState
End Sub


Private Sub cmdOkMaterials_Click()
    If txtQty.Value = 0 Then: Call ColorIt(txtQty, Timer1): Exit Sub
    If txtRate.Value = 0 Then: Call ColorIt(txtRate, Timer1): Exit Sub
    If txtAmount.Value = 0 Then: Call ColorIt(txtAmount, Timer1): Exit Sub

    Dim lst                                  As ReportRecord
    Dim totamount                            As Currency
    Dim i                                    As Integer
    Dim REC                                  As ReportRecord
    Dim ItemExist                            As Boolean
    ''FIND ITEM


    If ListEdit = True Then
        Set lst = lvQuotation.SelectedRows(0).Record
        prc_FillLines lst, lvQuotation.SelectedRows(0).Record(2).Value
        prc_UpdateSubTotal
        Set lst = Nothing
        MyPictureState
        Exit Sub
    End If

    For i = 0 To lvQuotation.Records.Count - 1
        If txtProductCode.Text = lvQuotation.Records(i).Item(1).Value Then
            Set lst = lvQuotation.Records(i)
            ItemExist = True
            Exit For
        End If
    Next


    ''OBJ EXISTS
    If Not lst Is Nothing Then
        If ItemExist = True Then
            If MsgBox("Item Exists In List." & vbCrLf & "Do you Want To Update?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Item Exists") = vbYes Then
                prc_FillLines lst, "Materials"
                Me.Refresh
            End If
        Else
            prc_FillLines lst, "Materials"
        End If
    Else
        ''ADD NEW LINE
        Set REC = lvQuotation.Records.Add
        REC.AddItem (lvQuotation.Records.Count)
        REC.AddItem txtProductCode.Text
        REC.AddItem QuotationType
        REC.AddItem cboProductList.Text
        REC.AddItem txtQty.Value
        REC.AddItem txtRate.Value
        REC.AddItem (txtAmount.Value)
        lvQuotation.Populate
    End If
    ''''''UPDATE AMOUNT
    prc_UpdateSubTotal
    Set lst = Nothing
    ''''''SET DEFAULT LINES
    MyPictureState

End Sub

Private Sub cmdOkServices_Click()
    'Validate it
    If txtSQty.Value = 0 Then: Call ColorIt(txtSQty, Timer1): Exit Sub
    If txtSRate.Value = 0 Then: Call ColorIt(txtSRate, Timer1): Exit Sub
    If txtSAmount.Value = 0 Then: Call ColorIt(txtSAmount, Timer1): Exit Sub
    Dim lst                                  As ReportRecord
    Dim totamount                            As Currency
    Dim i                                    As Integer
    Dim REC                                  As ReportRecord
    Dim ItemExist                            As Boolean

    If ListEdit = True Then                                   'If its Editing then Just Populate the Row
        Set lst = lvQuotation.SelectedRows(0).Record
        prc_FillLines lst, "Service"
        prc_UpdateSubTotal
        Set lst = Nothing
        MyPictureState
        Exit Sub
    End If
    For i = 0 To lvQuotation.Records.Count - 1                'Check for Item Existence
        If cboServiceList.Text = lvQuotation.Records(i).Item(1).Value Then
            Set lst = lvQuotation.Records(i)
            ItemExist = True
            Exit For
        End If
    Next
    If Not lst Is Nothing Then                                ''IF Row EXISTS THen
        If ItemExist = True Then                              'Ask User For Updates
            If MsgBox("Item Exists In List." & vbCrLf & "Do you Want To Update?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Item Exists") = vbYes Then
                prc_FillLines lst, "Service"
                Me.Refresh
            End If
        Else
            prc_FillLines lst, "Service"                      'If Doesn't Exist Then Update Again
        End If
    Else
        Set REC = lvQuotation.Records.Add                     ''INCASE NOTHING IS CONDITIONED THEN ADD NEW LINE
        REC.AddItem (lvQuotation.Records.Count)
        REC.AddItem cboServiceList.Text
        REC.AddItem QuotationType
        REC.AddItem cboServiceList.Text
        REC.AddItem txtSQty.Value
        REC.AddItem txtSRate.Value
        REC.AddItem txtSAmount.Value
        lvQuotation.Populate
    End If
    prc_UpdateSubTotal                                        ''''''UPDATE AMOUNT
    Set lst = Nothing
    MyPictureState                                            ''''''SET DEFAULT LINES
End Sub

Private Sub cmdOkVehicles_Click()
    If txtVNetRate.Value = 0 Then: Call ColorIt(txtVNetRate, Timer1): Exit Sub
    If txtVQty.Value = 0 Then: Call ColorIt(txtVQty, Timer1): Exit Sub
    If txtVTotalAmount.Value = 0 Then: Call ColorIt(txtVTotalAmount, Timer1): Exit Sub


    Dim lst                                  As ReportRecord
    Dim totamount                            As Currency
    Dim i                                    As Integer
    Dim REC                                  As ReportRecord
    Dim ItemExist                            As Boolean
    ''FIND ITEM


    If ListEdit = True Then
        Set lst = lvQuotationVehicles.SelectedRows(0).Record
        prc_FillLines lst, lvQuotationVehicles
        prc_UpdateSubTotal
        Set lst = Nothing
        MyPictureState
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
                prc_FillLines lst, "Vehicles"
                Me.Refresh
            End If
        Else
            prc_FillLines lst, "Vehicles"
        End If
    Else
        ''ADD NEW LINE
        Set REC = lvQuotationVehicles.Records.Add
        REC.AddItem (lvQuotationVehicles.Records.Count)
        REC.AddItem lblvCode.caption
        REC.AddItem lblvDescript
        REC.AddItem txtVQty.Value
        REC.AddItem txtVNetRate.Text
        REC.AddItem txtVdownpayment.Text
        REC.AddItem (cboVTerm.Text)
        REC.AddItem (txtVAOR.Value)
        REC.AddItem (txtVBalToFin.Text)
        REC.AddItem (txtvNetMonthlyMort.Text)

        lvQuotationVehicles.Populate
    End If
    ''''''UPDATE AMOUNT
    prc_UpdateSubTotal
    Set lst = Nothing
    ''''''SET DEFAULT LINES
    MyPictureState
End Sub

Private Sub cmdSave_Click()
    Dim i                                    As Integer
    Dim SQL                                  As String
    Dim temprs                               As ADODB.Recordset
    If QuotationID <= 0 Then
        SQL = "INSERT INTO CRIS_Quote_Header(ProspectID,  QuotationCode, QuotationDescription) " _
            & " VALUES(@ProspectID,  @QuotationCode, @QuotationDescription)" & vbCrLf & " SELECT @@IDENTITY"
    Else

        SQL = "  Update CRIS_Quote_Header " _
            & " SET QuotationCode=@QuotationCode, QuotationDescription=@QuotationDescription " _
            & " WHERE QuotationID=@QuotationID "

        SQL = " UPDATE CRIS_Quote_Header " _
            & " SET QuotationDescription=@QuotationDescription  " _
            & " WHERE   QuotationID=@QuotationID "
    End If
    SQL = Replace(SQL, "@QuotationID", QuotationID)
    SQL = Replace(SQL, "@QuotationCode", N2Str2Null(txtQuotationCode.Text))
    SQL = Replace(SQL, "@QuotationDescription", N2Str2Null(txtNotes.Text))
    SQL = Replace(SQL, "@ProspectID", ProspectID)
    SQL = Replace(SQL, "@ProfileType", N2Str2Null(ProfileType))



    Set temprs = gconDMIS.Execute(SQL)
    Set temprs = temprs.NextRecordset
    If Not temprs Is Nothing Then
        QuotationID = temprs.Collect(0)
    End If

    gconDMIS.Execute ("DELETE FROM CRIS_Quote_Details WHERE QuotationCode=" & N2Str2Null(txtQuotationCode.Text))
    For i = 0 To lvQuotation.Records.Count - 1
        With lvQuotation.Records(i)
            SQL = "INSERT INTO CRIS_Quote_Details " & _
                "   (QuotationCode, EntryCode, Price,Qty,QuotationType) " & _
                  "VALUES (@QuotationCode, @EntryCode, @Price,@QTY,@QuotationType) "
            SQL = Replace(SQL, "@QuotationCode", N2Str2Null(txtQuotationCode.Text))
            SQL = Replace(SQL, "@EntryCode", N2Str2Null(.Item(1).Value))
            SQL = Replace(SQL, "@QuotationType", N2Str2Null(Left(.Item(2).Value, 1)))
            SQL = Replace(SQL, "@QTY", .Item(4).Value)
            SQL = Replace(SQL, "@Price", .Item(5).Value)
            gconDMIS.Execute SQL
        End With
    Next
    For i = 0 To lvQuotationVehicles.Records.Count - 1
        With lvQuotationVehicles.Records(i)
            SQL = "INSERT INTO CRIS_Quote_Details " & _
                "   (QuotationCode, EntryCode, Price,Qty,QuotationType, Downpayment, Terms,AOR, BalToFin, NetMonthlyAmort) " & _
                  "VALUES (@QuotationCode, @EntryCode, @Price,@QTY,@QuotationType,@DownPayment, @Terms, @AOR , @BalToFin , @NetMonthlyAmort ) "

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
    SQL = Replace(SQL, "@TotalAmount", txtTotalAmount.Value)
    gconDMIS.Execute SQL
    MessagePop RecSave, "Record Saved", " Record Saved "
    txtQuotationCode.Enabled = False
    gconDMIS.Execute ("Update CRIS_PROSPECTS SET LogQuote=getdate() Where ProspectID=" & ProspectID)
    
    RaiseEvent ChangeList(True)

End Sub

Public Sub FillListData(objx As Object, oRS As Recordset, ShowingIndex As String)

    objx.Clear
    While Not oRS.EOF
        objx.AddItem (oRS.Fields(ShowingIndex))
        objx.ItemData(objx.NewIndex) = oRS.Collect(0)
        oRS.MoveNext
    Wend
    If objx.ListCount > 0 Then
        If TypeName(objx) = "ComboBox" Then
            objx.ListIndex = 0
        End If
    End If
End Sub



Private Sub Form_Load()
    Dim lvwidth                              As Long
    
    CenterPicture picServiceDetails
    CenterPicture picProductDetails
    CenterPicture picVehiclesDetails
    Call SendMessage(cboJobCategory.hwnd, CB_SETDROPPEDWIDTH, 300, 0)
    Call SendMessage(cboVehicles.hwnd, CB_SETDROPPEDWIDTH, 300, 0)
    Call SendMessage(cboServiceList.hwnd, CB_SETDROPPEDWIDTH, 400, 0)

    Set RsParts_N_Mat = gconDMIS.Execute("SELECT  " & _
                                        "ID,  " & _
                                        "STOCKNO as Code,  " & _
                                        "ISNULL(SRP,0) as Rate,  " & _
                                        "STOCKDESC as Particulars ,  " & _
                                        "TYPE , " & _
                                        "ISNULL(ONHAND, 0)  as OnHand from  " & _
                                        "PMIS_STOCKMAS Order by Type , StockDesc")
    Set RsJobs = gconDMIS.Execute("SELECT      " & _
                                 "CJ.ID, " & _
                                 "CJ.JCode as Code,  " & _
                                 "CJC.[desc] AS Category,  " & _
                                 "ISNULL(CJC.FlatRate * CJ.std_mhrs,0) AS Rate,  " & _
                                 "CJ.Desc1 as Particulars " & _
                                 "FROM          " & _
                               " CSMS_JobCategory CJC " & _
                                 "INNER JOIN " & _
                                 "CSMS_Jobs CJ " & _
                                 "ON      " & _
                                 "CJC.Jcat = CJ.JCat  ")


    Set RsVehicles = gconDMIS.Execute("SELECT ID,CODE, DESCRIPT as Particulars, MODEL, MAKE, YEarModel  FROM CRIS_MRRINV")
    If (RsParts_N_Mat.EOF Or RsParts_N_Mat.BOF) = True Then
        MessagePop InfoVoid, "InSuffcient Record", "There Are No Enlisted Master File For Parts or Materials. Please Add Few "
    ElseIf (RsJobs.EOF Or RsJobs.BOF) = True Then
        MessagePop InfoVoid, "InSuffcient Record", "There Are No Enlisted Master File For Jobs. Please Add Few "
    ElseIf (RsVehicles.EOF Or RsVehicles.BOF) = True Then
        MessagePop InfoVoid, "InSuffcient Record", "There Are No Enlisted Master File For Vehicles. Please Add Few "
    End If
    FillListData cboJobCategory, gconDMIS.Execute("SELECT ID, [desc] AS Category FROM CSMS_JobCategory "), "Category"
    FillListData cboVehicles, RsVehicles, "Particulars"
    lvwidth = lvQuotation.Width
    With lvQuotation
        .Columns.Add 0, "Item", 30, False
        .Columns.Add 1, "CODE", 90, False
        .Columns.Add 2, "Type", 100, False
        .Columns.Add 3, "Description", 250, False
        .Columns.Add 4, "QTY", 35, False
        .Columns.Add 5, "Rate", 50, False
        .Columns.Add 6, "Amount", 120, False
        .GroupsOrder.Add .Columns(2)
        .Columns(2).Visible = False
    End With
    With lvQuotationVehicles
        .Columns.Add 0, "Item", 30, False
        .Columns.Add 1, "CODE", 80, False
        .Columns.Add 2, "Description", 165, False
        .Columns.Add 3, "QTY", 35, False
        .Columns.Add 4, "Net Price", 120, False
    End With


    With lvSearch
        .Columns.Add 1, "Date", 80, False
        .Columns.Add 2, "Particulars", 165, False
    End With
    With lvQuotation
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots  ' xtpGridNoLines
        .PaintManager.GroupRowTextBold = True                 ' = vbWhite
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
    End With
    With lvQuotationVehicles
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots  ' xtpGridNoLines
        .PaintManager.GroupRowTextBold = True                 ' = vbWhite
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





End Sub


'Private Sub cboProductCode_Click()
'   ' cboProductList.ListIndex = SelectCombo(cboProductCode, cboProductCode.ItemData(cboProductCode.ListIndex), True)
'End Sub
'Private Sub cboProductList_Click()
'    '
'End Sub


Private Sub Form_Unload(Cancel As Integer)
    QuotationID = 0
End Sub


Sub LabelIt()
    Dim temprs                               As ADODB.Recordset
     Set temprs = gconDMIS.Execute("select * from   CRIS_vW_AllProfile where Profileid=" & ProfileID & " and ProfileTYpe =" & N2Str2Null(ProfileType))
     
    If Not (temprs.EOF Or temprs.BOF) Then
        lblCustomerName(0).caption = temprs!ProfileName
        lblAccountName(0).caption = temprs!AcctName
        lblAddress(0).caption = temprs!Address
        lblContactNo(0).caption = temprs!Phone
        lblEmail(0).caption = temprs!Email
        If temprs!ProfileType = "CC" Then
            capCustomerDetails.caption = "::::Company Customer"
        ElseIf temprs!ProfileType = "CP" Then
            capCustomerDetails.caption = "::::Personal Customer"
        ElseIf temprs!ProfileType = "PP" Then
            capCustomerDetails.caption = "::::Personal Prospects"
        ElseIf temprs!ProfileType = "PC" Then
            capCustomerDetails.caption = "::::Company Prospects"
        End If

        lblCustomerName(0).caption = temprs!ProfileName
        lblAccountName(0).caption = temprs!AcctName
        lblAddress(0).caption = temprs!Address
        lblContactNo(0).caption = temprs!Phone
        lblEmail(0).caption = temprs!Email
        If temprs!ProfileType = "CC" Then
            capCustomerDetails.caption = "::::Company Customer"
        ElseIf temprs!ProfileType = "CP" Then
            capCustomerDetails.caption = "::::Personal Customer"
        ElseIf temprs!ProfileType = "PP" Then
            capCustomerDetails.caption = "::::Personal Prospects"
        ElseIf temprs!ProfileType = "PC" Then
            capCustomerDetails.caption = "::::Company Prospects"
        End If
    End If
    Set temprs = Nothing
End Sub

Sub LabelJobDescriptions()
    If cboJobCodes.ListIndex = -1 Then: Exit Sub
    Dim temprs                               As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT JModel, OPCODE,  flatrate, std_mhrs From CSMS_Jobs Where ID=" & cboJobCodes.ItemData(cboJobCodes.ListIndex))
    If Not (temprs.EOF Or temprs.BOF) Then
        lblJcode.caption = cboJobCodes.Text
        lblJDescript.caption = cboServiceList.Text
        lblJFlatRate.caption = temprs!flatrate
        lblJModel.caption = temprs!JModel
        lblJOpCode.caption = temprs!OPCODE
        lblJStd_MHRS.caption = temprs!std_mhrs


    End If
End Sub

Private Sub lvQuotation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If lvQuotation.SelectedRows.Count = 0 Then: Exit Sub
        If MsgBox("Confirm Your Action ", vbYesNo Or vbInformation Or vbDefaultButton1, "Confirm Deletion") = vbYes Then
            Connect
            gconDMIS.Execute ("Delete from CRIS_Quote_Details where QuotationCode=" & N2Str2Null(txtQuotationCode.Text) & " and EntryCode=" & N2Str2Null(lvQuotation.SelectedRows.Row(0).Record(1).Value))
            prc_FillDetails
            'txtSearch.SetFocus
        End If
    End If
End Sub

Private Sub lvQuotation_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then: Exit Sub
    Select Case Row.Record(2).Value
        Case "Materials"
            cap2.caption = "Add Materials"
            cmdAddMaterials_Click
            ListEdit = True
            txtProductCode.Enabled = False
            cboProductList.Enabled = False
            cboProductList.ListIndex = SelectCombo(cboProductList, Row.Record(3).Value, False)
            txtQty.Value = Row.Record(4).Value
            txtRate.Value = Row.Record(5).Value
            txtAmount.Value = Row.Record(6).Value
        Case "Services"
            cmdAddService_Click
            ListEdit = True
            cboJobCategory.Enabled = False
            cboServiceList.Enabled = False
            cboJobCodes.Enabled = False
            cboServiceList.Text = Row.Record(3).Value
            txtSQty.Value = Row.Record(4).Value
            txtSRate.Value = Row.Record(5).Value
            txtSAmount.Value = Row.Record(6).Value
        Case "Parts"
            cap2.caption = "Add Parts"
            cmdAddParts_Click
            ListEdit = True
            txtProductCode.Enabled = False
            cboProductList.Enabled = False
            txtProductCode.Text = Row.Record(1).Value
            cboProductList.Text = Row.Record(3).Value
            txtQty.Value = Row.Record(4).Value
            txtRate.Value = Row.Record(5).Value
            txtAmount.Value = Row.Record(6).Value
        Case "Vehicles"
            cmdAddVehicles_Click
    End Select
    '    txtSearch.Enabled = False
    '    lstProd.Enabled = False
    '    cmdCancelLine.Enabled = True
    '    txtSearch.Tag = Row.Record(1).Value
    '    txtSearch.Text = Row.Record(3).Value
    '    txtQty.Value = Row.Record(4).Value
    '    txtRate.Value = Row.Record(5).Value
    '    txtAmount.Value = Row.Record(6).Value
    '    txtQty.SetFocus
End Sub

Sub MyPictureState()
    picTop.Enabled = True
    picDetails.Enabled = True
    picNames.Enabled = True
    picProductDetails.Visible = False
    picServiceDetails.Visible = False
    picVehiclesDetails.Visible = False
End Sub

Friend Sub NewQuotation(xProfileID As Long, xProfileType As String, xAcctName As String, xProspectID As Long)
    ProfileID = xProfileID
    ProfileType = xProfileType
    AcctName = xAcctName
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
    Dim temprs                               As ADODB.Recordset
    Dim SQL                                  As String
    Dim FLD                                  As Field
    Dim j                                    As Long
    Dim rec1                                 As XtremeReportControl.ReportRecord
    Dim rec2                                 As XtremeReportControl.ReportRecord


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
    lvQuotation.Records.DeleteAll
    While Not temprs.EOF
        j = j + 1
        Set rec1 = lvQuotation.Records.Add
        rec1.AddItem j
        For Each FLD In temprs.Fields
            rec1.AddItem (Trim(FLD.Value))
        Next
        temprs.MoveNext
    Wend
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
    SQL = "Select Code, [Description], Qty, Price,  DownPayment, Terms, AOR, BalToFin, NetMonthlyAmort From CRIS_Vw_QuotationDetails  Where Type='Vehicles' and QuotationCode=" & N2Str2Null(txtQuotationCode.Text)
    Set temprs = gconDMIS.Execute(SQL)
    j = 0
    While Not temprs.EOF
        j = j + 1
        Set rec2 = lvQuotationVehicles.Records.Add
        rec2.AddItem j
        For Each FLD In temprs.Fields
            If FLD.Type = adDouble Then
                rec2.AddItem (FormatCurrency(FLD.Value, 2, vbTrue, vbTrue, vbTrue))
            Else
                rec2.AddItem (Trim(FLD.Value))
            End If

        Next
        temprs.MoveNext
    Wend


    lvQuotation.Populate
    lvQuotationVehicles.Populate
    Set FLD = Nothing
    Set rec1 = Nothing
    Set rec2 = Nothing
    Set temprs = Nothing
    'prc_UpdateSubTotal


End Sub

Private Sub prc_FillHeader()
    If QuotationID <= 0 Then: Exit Sub


    Dim oRsx                                 As ADODB.Recordset

    Set oRsx = GetRS("Select * from CRIS_Quote_Header Where QuotationID=" & QuotationID)

    If Not oRsx.EOF Or oRsx.BOF Then
        txtQuotationCode.Text = oRsx.Fields("QuotationCode")
        txtNotes.Text = Null2String(oRsx.Fields("QuotationDescription"))
        txtQuotationCode.Enabled = False
        oRsx.MoveNext
    End If
    Set oRsx = Nothing
End Sub

Private Sub prc_FillLines(lst As ReportRecord, QuotationType As String)
    Select Case QuotationType
        Case "Vehicles"
            With lst
                .Item(4).Value = txtVNetRate.Value
                .Item(5).Value = txtVdownpayment.Value
                .Item(6).Value = cboVTerm.Text
                .Item(7).Value = txtVAOR.Value
                .Item(8).Value = txtVBalToFin.Value
                .Item(9).Value = txtvNetMonthlyMort.Value

                lvQuotationVehicles.Populate
            End With
        Case "Service"
            With lst
                .Item(4).Value = txtSQty.Value
                .Item(5).Value = txtSRate.Value
                .Item(6).Value = (txtSQty.Value * txtSRate.Value)
            End With
            lvQuotation.Populate
        Case "Parts"
            With lst
                .Item(4).Value = txtQty.Value
                .Item(5).Value = txtRate.Value
                .Item(6).Value = (txtQty.Value * txtRate.Value)
            End With
            lvQuotation.Populate
        Case "Materials"
            With lst
                .Item(4).Value = txtQty.Value
                .Item(5).Value = txtRate.Value
                .Item(6).Value = (txtQty.Value * txtRate.Value)
            End With
            lvQuotation.Populate
    End Select
End Sub

Private Sub prc_UpdateSubTotal()
    Dim i                                    As Integer
    Dim totamount                            As Currency
    For i = 0 To lvQuotation.Records.Count - 1
        ' totamount = totamount + CCur(lvQuotation.Records(i).ITEM(6).Value)
    Next
    txtTotalAmount.Value = totamount
End Sub

Sub ShowForm(hwnd As Long)
    Dim cntl                                 As Control
    For Each cntl In Me.ControlS
        If TypeOf cntl Is PictureBox Then
            If cntl.hwnd = hwnd Then
                cntl.Enabled = True
                cntl.Visible = True
                cntl.ZOrder 0
            Else
                cntl.Enabled = False
                cntl.ZOrder 1
            End If
        End If
    Next
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
    Dim cntrl                                As Control
    For Each cntrl In Me.ControlS
        If TypeOf cntrl Is TextBox Or _
           TypeOf cntrl Is ComboBox Or _
           TypeOf cntrl Is fpLongInteger Or _
           TypeOf cntrl Is fpDoubleSingle Or _
           TypeOf cntrl Is fpText Or _
           TypeOf cntrl Is fpCurrency Then
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

Private Sub txtProductCode_LostFocus()
    RsParts_N_Mat.Filter = "CODE='" & txtProductCode.Text & "'"
    If Not (RsParts_N_Mat.EOF Or RsParts_N_Mat.BOF) Then
        cboProductList.Text = RsParts_N_Mat.Fields("Particulars").Value
    Else
        txtProductCode.Text = vbNullString
    End If

End Sub

Private Sub txtQty_Change()
    If txtQty.Value > 0 Then
        txtAmount.Value = txtRate.Value * txtQty.Value
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: SendKeys ("{TAB}")
End Sub

Private Sub txtQuotationCode_Change()
    'if
    '   cmdSave.Enabled = True
End Sub

Private Sub txtQuotationCode_Validate(Cancel As Boolean)
    If Len(Trim(txtQuotationCode.Text)) < 6 Then
        Call ColorIt(txtQuotationCode, Timer1)
        Cancel = True
    Else
        If gconDMIS.Execute("Select COUNT(*) FROM CRIS_Quote_Header WHERE QuotationCode=" & N2Str2Null(txtQuotationCode.Text)).Fields(0).Value > 0 Then
            Call ColorIt(txtQuotationCode, Timer1)
            MessagePop RecSaveError, "Duplicated Entry", "Same Code Exist. Please Use Another Code"
            Cancel = True
        End If
    End If
End Sub

Private Sub txtRate_Change()
    txtAmount.Value = txtRate.Value * txtQty.Value
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then: txtQty.SetFocus
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: SendKeys ("{TAB}")
End Sub

Private Sub txtSQty_Change()
    txtSAmount.Value = txtSRate.Value * txtSQty.Value
End Sub

'Private Sub txtsearch_ButtonHit(Button As Integer, NewIndex As Integer)
'    If Button = 5 Then
'        If txtsearch.Tag <> 0 Or txtsearch.Tag <> vbNullString Then
'            EntryProduct.MainID = txtsearch.Tag
'        End If
'        EntryProduct.Show vbModal
'        If EntryProduct.ValueChanges = True Then
'            mEdit = True
'            txtsearch.Text = vbNullString
'            lstProd.Clear
'            prc_FillData
'        End If
'    End If
'End Sub

Private Sub txtVNetRate_LostFocus()
    txtVTotalAmount.Value = txtVQty.Value * txtVNetRate.Value
End Sub

Private Sub txtVQty_Change()
    txtVTotalAmount.Value = txtVQty.Value * txtVNetRate.Value
End Sub

