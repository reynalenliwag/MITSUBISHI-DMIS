VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Begin VB.Form frmAccMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password File Maintenance"
   ClientHeight    =   3345
   ClientLeft      =   5415
   ClientTop       =   2445
   ClientWidth     =   3870
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Accmaintenance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3345
   ScaleWidth      =   3870
   Begin VB.PictureBox picCOND 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9195
      Left            =   10200
      ScaleHeight     =   9165
      ScaleWidth      =   13500
      TabIndex        =   9
      Top             =   90
      Width           =   13530
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4965
         Left            =   780
         ScaleHeight     =   4935
         ScaleWidth      =   5355
         TabIndex        =   26
         Top             =   630
         Width           =   5385
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   10
            Left            =   30
            ScaleHeight     =   255
            ScaleWidth      =   5265
            TabIndex        =   86
            Top             =   4620
            Width           =   5295
            Begin VB.OptionButton optACC 
               Caption         =   "HEAVY ELECTRIC LOAD"
               Height          =   240
               Index           =   1
               Left            =   2535
               MouseIcon       =   "Accmaintenance.frx":1472
               MousePointer    =   99  'Custom
               TabIndex        =   88
               Top             =   0
               Width           =   2475
            End
            Begin VB.OptionButton optACC 
               Caption         =   "A/C ON"
               Height          =   240
               Index           =   0
               Left            =   1240
               MouseIcon       =   "Accmaintenance.frx":15C4
               MousePointer    =   99  'Custom
               TabIndex        =   87
               Top             =   0
               Value           =   -1  'True
               Width           =   1395
            End
            Begin VB.Label lblcap 
               BackColor       =   &H00FCE2CF&
               Caption         =   " ACCESSORIES"
               Height          =   270
               Index           =   26
               Left            =   0
               TabIndex        =   89
               Top             =   0
               Width           =   1260
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   30
            ScaleHeight     =   225
            ScaleWidth      =   5265
            TabIndex        =   82
            Top             =   4347
            Width           =   5295
            Begin VB.OptionButton optOCC 
               Caption         =   "INTERMITTENT"
               Height          =   255
               Index           =   1
               Left            =   2535
               MouseIcon       =   "Accmaintenance.frx":1716
               MousePointer    =   99  'Custom
               TabIndex        =   84
               Top             =   0
               Width           =   1425
            End
            Begin VB.OptionButton optOCC 
               Caption         =   "CONSISTENT"
               Height          =   255
               Index           =   0
               Left            =   1240
               MouseIcon       =   "Accmaintenance.frx":1868
               MousePointer    =   99  'Custom
               TabIndex        =   83
               Top             =   0
               Value           =   -1  'True
               Width           =   1245
            End
            Begin VB.Label lblcap 
               BackColor       =   &H00FCE2CF&
               Caption         =   " OCCURENCE"
               Height          =   270
               Index           =   27
               Left            =   0
               MouseIcon       =   "Accmaintenance.frx":19BA
               MousePointer    =   99  'Custom
               TabIndex        =   85
               Top             =   0
               Width           =   1260
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   555
            Index           =   7
            Left            =   30
            ScaleHeight     =   525
            ScaleWidth      =   5265
            TabIndex        =   75
            Top             =   3774
            Width           =   5295
            Begin VB.OptionButton optACT 
               Caption         =   "IDLING"
               Height          =   240
               Index           =   1
               Left            =   2535
               MouseIcon       =   "Accmaintenance.frx":1B0C
               MousePointer    =   99  'Custom
               TabIndex        =   80
               Top             =   0
               Width           =   1065
            End
            Begin VB.OptionButton optACT 
               Caption         =   "DECELERATING"
               Height          =   240
               Index           =   4
               Left            =   3495
               MouseIcon       =   "Accmaintenance.frx":1C5E
               MousePointer    =   99  'Custom
               TabIndex        =   79
               Top             =   270
               Width           =   1695
            End
            Begin VB.OptionButton optACT 
               Caption         =   "ACCELERATING"
               Height          =   240
               Index           =   3
               Left            =   1995
               MouseIcon       =   "Accmaintenance.frx":1DB0
               MousePointer    =   99  'Custom
               TabIndex        =   78
               Top             =   270
               Width           =   1425
            End
            Begin VB.OptionButton optACT 
               Caption         =   "CRUISING"
               Height          =   240
               Index           =   2
               Left            =   3855
               MouseIcon       =   "Accmaintenance.frx":1F02
               MousePointer    =   99  'Custom
               TabIndex        =   77
               Top             =   0
               Width           =   1305
            End
            Begin VB.OptionButton optACT 
               Caption         =   "CRANKING"
               Height          =   240
               Index           =   0
               Left            =   1215
               MouseIcon       =   "Accmaintenance.frx":2054
               MousePointer    =   99  'Custom
               TabIndex        =   76
               Top             =   0
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.Label lblcap 
               BackColor       =   &H00FCE2CF&
               Caption         =   " ACTION"
               Height          =   540
               Index           =   28
               Left            =   0
               MouseIcon       =   "Accmaintenance.frx":21A6
               MousePointer    =   99  'Custom
               TabIndex        =   81
               Top             =   0
               Width           =   1260
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   525
            Index           =   8
            Left            =   30
            ScaleHeight     =   495
            ScaleWidth      =   5265
            TabIndex        =   69
            Top             =   3231
            Width           =   5295
            Begin VB.OptionButton optLOC 
               Caption         =   "UPHILL"
               Height          =   255
               Index           =   1
               Left            =   2535
               MouseIcon       =   "Accmaintenance.frx":22F8
               MousePointer    =   99  'Custom
               TabIndex        =   73
               Top             =   0
               Width           =   1125
            End
            Begin VB.OptionButton optLOC 
               Caption         =   "DOWNHILL"
               Height          =   255
               Index           =   2
               Left            =   3855
               MouseIcon       =   "Accmaintenance.frx":244A
               MousePointer    =   99  'Custom
               TabIndex        =   72
               Top             =   0
               Width           =   1635
            End
            Begin VB.OptionButton optLOC 
               Caption         =   "HIGHWAY"
               Height          =   255
               Index           =   0
               Left            =   1245
               MouseIcon       =   "Accmaintenance.frx":259C
               MousePointer    =   99  'Custom
               TabIndex        =   71
               Top             =   0
               Value           =   -1  'True
               Width           =   1125
            End
            Begin VB.OptionButton optLOC 
               Caption         =   "STOP AND GO TRAFFIC"
               Height          =   255
               Index           =   3
               Left            =   1240
               MouseIcon       =   "Accmaintenance.frx":26EE
               MousePointer    =   99  'Custom
               TabIndex        =   70
               Top             =   270
               Width           =   2205
            End
            Begin VB.Label lblcap 
               BackColor       =   &H00FCE2CF&
               Caption         =   " LOCATION"
               Height          =   510
               Index           =   29
               Left            =   0
               MouseIcon       =   "Accmaintenance.frx":2840
               MousePointer    =   99  'Custom
               TabIndex        =   74
               Top             =   15
               Width           =   1260
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   555
            Index           =   9
            Left            =   30
            ScaleHeight     =   525
            ScaleWidth      =   5265
            TabIndex        =   63
            Top             =   2658
            Width           =   5295
            Begin VB.OptionButton optROAD 
               Caption         =   "INPAVED"
               Height          =   240
               Index           =   1
               Left            =   2520
               MouseIcon       =   "Accmaintenance.frx":2992
               MousePointer    =   99  'Custom
               TabIndex        =   67
               Top             =   0
               Width           =   975
            End
            Begin VB.OptionButton optROAD 
               Caption         =   "MUDDY"
               Height          =   240
               Index           =   3
               Left            =   1240
               MouseIcon       =   "Accmaintenance.frx":2AE4
               MousePointer    =   99  'Custom
               TabIndex        =   66
               Top             =   270
               Width           =   1005
            End
            Begin VB.OptionButton optROAD 
               Caption         =   "ROCKY"
               Height          =   240
               Index           =   2
               Left            =   3855
               MouseIcon       =   "Accmaintenance.frx":2C36
               MousePointer    =   99  'Custom
               TabIndex        =   65
               Top             =   0
               Width           =   945
            End
            Begin VB.OptionButton optROAD 
               Caption         =   "PAVED"
               Height          =   240
               Index           =   0
               Left            =   1240
               MouseIcon       =   "Accmaintenance.frx":2D88
               MousePointer    =   99  'Custom
               TabIndex        =   64
               Top             =   0
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.Label lblcap 
               BackColor       =   &H00FCE2CF&
               Caption         =   " ROAD"
               Height          =   540
               Index           =   31
               Left            =   0
               TabIndex        =   68
               Top             =   0
               Width           =   1260
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   555
            Index           =   2
            Left            =   30
            ScaleHeight     =   525
            ScaleWidth      =   5265
            TabIndex        =   54
            Top             =   2085
            Width           =   5295
            Begin VB.OptionButton optAT 
               Caption         =   "R"
               Height          =   240
               Index           =   5
               Left            =   2535
               MouseIcon       =   "Accmaintenance.frx":2EDA
               MousePointer    =   99  'Custom
               TabIndex        =   61
               Top             =   260
               Width           =   1065
            End
            Begin VB.OptionButton optAT 
               Caption         =   "2"
               Height          =   240
               Index           =   1
               Left            =   2535
               MouseIcon       =   "Accmaintenance.frx":302C
               MousePointer    =   99  'Custom
               TabIndex        =   60
               Top             =   0
               Width           =   1065
            End
            Begin VB.OptionButton optAT 
               Caption         =   "OVERDRIVE"
               Height          =   240
               Index           =   6
               Left            =   3855
               MouseIcon       =   "Accmaintenance.frx":317E
               MousePointer    =   99  'Custom
               TabIndex        =   59
               Top             =   260
               Width           =   1755
            End
            Begin VB.OptionButton optAT 
               Caption         =   "N"
               Height          =   240
               Index           =   4
               Left            =   1230
               MouseIcon       =   "Accmaintenance.frx":32D0
               MousePointer    =   99  'Custom
               TabIndex        =   58
               Top             =   260
               Width           =   1065
            End
            Begin VB.OptionButton optAT 
               Caption         =   "D"
               Height          =   240
               Index           =   3
               Left            =   4500
               MouseIcon       =   "Accmaintenance.frx":3422
               MousePointer    =   99  'Custom
               TabIndex        =   57
               Top             =   0
               Width           =   705
            End
            Begin VB.OptionButton optAT 
               Caption         =   "3"
               Height          =   240
               Index           =   2
               Left            =   3855
               MouseIcon       =   "Accmaintenance.frx":3574
               MousePointer    =   99  'Custom
               TabIndex        =   56
               Top             =   0
               Width           =   1065
            End
            Begin VB.OptionButton optAT 
               Caption         =   "L"
               Height          =   240
               Index           =   0
               Left            =   1230
               MouseIcon       =   "Accmaintenance.frx":36C6
               MousePointer    =   99  'Custom
               TabIndex        =   55
               Top             =   0
               Value           =   -1  'True
               Width           =   1065
            End
            Begin VB.Label lblcap 
               BackColor       =   &H00FCE2CF&
               Caption         =   " A/T"
               Height          =   540
               Index           =   30
               Left            =   0
               TabIndex        =   62
               Top             =   15
               Width           =   1260
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   555
            Index           =   3
            Left            =   30
            ScaleHeight     =   525
            ScaleWidth      =   5265
            TabIndex        =   45
            Top             =   1512
            Width           =   5295
            Begin VB.OptionButton optMT 
               Caption         =   "5th"
               Height          =   240
               Index           =   5
               Left            =   2535
               MouseIcon       =   "Accmaintenance.frx":3818
               MousePointer    =   99  'Custom
               TabIndex        =   52
               Top             =   270
               Width           =   885
            End
            Begin VB.OptionButton optMT 
               Caption         =   "2ND"
               Height          =   240
               Index           =   1
               Left            =   2535
               MouseIcon       =   "Accmaintenance.frx":396A
               MousePointer    =   99  'Custom
               TabIndex        =   51
               Top             =   0
               Width           =   885
            End
            Begin VB.OptionButton optMT 
               Caption         =   "REVERSE"
               Height          =   240
               Index           =   6
               Left            =   3855
               MouseIcon       =   "Accmaintenance.frx":3ABC
               MousePointer    =   99  'Custom
               TabIndex        =   50
               Top             =   270
               Width           =   1065
            End
            Begin VB.OptionButton optMT 
               Caption         =   "5th"
               Height          =   240
               Index           =   4
               Left            =   1260
               MouseIcon       =   "Accmaintenance.frx":3C0E
               MousePointer    =   99  'Custom
               TabIndex        =   49
               Top             =   270
               Width           =   1065
            End
            Begin VB.OptionButton optMT 
               Caption         =   "4th"
               Height          =   240
               Index           =   3
               Left            =   4530
               MouseIcon       =   "Accmaintenance.frx":3D60
               MousePointer    =   99  'Custom
               TabIndex        =   48
               Top             =   0
               Width           =   1065
            End
            Begin VB.OptionButton optMT 
               Caption         =   "3rd"
               Height          =   240
               Index           =   2
               Left            =   3855
               MouseIcon       =   "Accmaintenance.frx":3EB2
               MousePointer    =   99  'Custom
               TabIndex        =   47
               Top             =   0
               Width           =   1065
            End
            Begin VB.OptionButton optMT 
               Caption         =   "1ST"
               Height          =   240
               Index           =   0
               Left            =   1260
               MouseIcon       =   "Accmaintenance.frx":4004
               MousePointer    =   99  'Custom
               TabIndex        =   46
               Top             =   0
               Value           =   -1  'True
               Width           =   1065
            End
            Begin VB.Label lblcap 
               BackColor       =   &H00FCE2CF&
               Caption         =   " M/T"
               Height          =   540
               Index           =   35
               Left            =   0
               TabIndex        =   53
               Top             =   0
               Width           =   1260
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   4
            Left            =   30
            ScaleHeight     =   255
            ScaleWidth      =   5265
            TabIndex        =   41
            Top             =   1209
            Width           =   5295
            Begin VB.OptionButton optPOS 
               Caption         =   "2WD"
               Height          =   240
               Index           =   1
               Left            =   2535
               MouseIcon       =   "Accmaintenance.frx":4156
               MousePointer    =   99  'Custom
               TabIndex        =   43
               Top             =   0
               Width           =   1240
            End
            Begin VB.OptionButton optPOS 
               Caption         =   "4WD"
               Height          =   240
               Index           =   0
               Left            =   1260
               MouseIcon       =   "Accmaintenance.frx":42A8
               MousePointer    =   99  'Custom
               TabIndex        =   42
               Top             =   0
               Value           =   -1  'True
               Width           =   1240
            End
            Begin VB.Label lblcap 
               BackColor       =   &H00FCE2CF&
               Caption         =   " SHIFT POSITION"
               Height          =   270
               Index           =   34
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Width           =   1260
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   5
            Left            =   30
            ScaleHeight     =   255
            ScaleWidth      =   5265
            TabIndex        =   37
            Top             =   906
            Width           =   5295
            Begin VB.OptionButton optSHIF 
               Caption         =   "FAST"
               Height          =   240
               Index           =   1
               Left            =   2535
               MouseIcon       =   "Accmaintenance.frx":43FA
               MousePointer    =   99  'Custom
               TabIndex        =   39
               Top             =   15
               Width           =   1245
            End
            Begin VB.OptionButton optSHIF 
               Caption         =   "NORMAL"
               Height          =   240
               Index           =   0
               Left            =   1260
               MouseIcon       =   "Accmaintenance.frx":454C
               MousePointer    =   99  'Custom
               TabIndex        =   38
               Top             =   15
               Value           =   -1  'True
               Width           =   1245
            End
            Begin VB.Label lblcap 
               BackColor       =   &H00FCE2CF&
               Caption         =   " SHIFTING"
               Height          =   240
               Index           =   33
               Left            =   0
               TabIndex        =   40
               Top             =   15
               Width           =   1260
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   6
            Left            =   30
            ScaleHeight     =   255
            ScaleWidth      =   5265
            TabIndex        =   32
            Top             =   603
            Width           =   5295
            Begin VB.OptionButton optWEA 
               Caption         =   "WARM"
               Height          =   240
               Index           =   1
               Left            =   1260
               MouseIcon       =   "Accmaintenance.frx":469E
               MousePointer    =   99  'Custom
               TabIndex        =   35
               Top             =   0
               Width           =   1245
            End
            Begin VB.OptionButton optWEA 
               Caption         =   "ALL TEMP"
               Height          =   240
               Index           =   2
               Left            =   3810
               MouseIcon       =   "Accmaintenance.frx":47F0
               MousePointer    =   99  'Custom
               TabIndex        =   34
               Top             =   0
               Width           =   1245
            End
            Begin VB.OptionButton optWEA 
               Caption         =   "COLD"
               Height          =   240
               Index           =   0
               Left            =   2535
               MouseIcon       =   "Accmaintenance.frx":4942
               MousePointer    =   99  'Custom
               TabIndex        =   33
               Top             =   0
               Value           =   -1  'True
               Width           =   1245
            End
            Begin VB.Label lblcap 
               Appearance      =   0  'Flat
               BackColor       =   &H00FCE2CF&
               Caption         =   " WEATHER"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   32
               Left            =   0
               TabIndex        =   36
               Top             =   0
               Width           =   1260
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   30
            ScaleHeight     =   255
            ScaleWidth      =   5265
            TabIndex        =   27
            Top             =   300
            Width           =   5295
            Begin VB.OptionButton optENG 
               Caption         =   "COLD"
               Height          =   240
               Index           =   1
               Left            =   2535
               MouseIcon       =   "Accmaintenance.frx":4A94
               MousePointer    =   99  'Custom
               TabIndex        =   30
               Top             =   0
               Width           =   1245
            End
            Begin VB.OptionButton optENG 
               Caption         =   "ALL TEMP"
               Height          =   240
               Index           =   2
               Left            =   3810
               MouseIcon       =   "Accmaintenance.frx":4BE6
               MousePointer    =   99  'Custom
               TabIndex        =   29
               Top             =   0
               Width           =   1245
            End
            Begin VB.OptionButton optENG 
               Caption         =   "HOT"
               Height          =   240
               Index           =   0
               Left            =   1260
               MouseIcon       =   "Accmaintenance.frx":4D38
               MousePointer    =   99  'Custom
               TabIndex        =   28
               Top             =   0
               Value           =   -1  'True
               Width           =   1245
            End
            Begin VB.Label lblcap 
               BackColor       =   &H00FCE2CF&
               Caption         =   " ENGINE TEMP"
               Height          =   240
               Index           =   36
               Left            =   0
               TabIndex        =   31
               Top             =   0
               Width           =   1260
            End
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   285
            Left            =   0
            TabIndex        =   90
            Top             =   0
            Width           =   6375
            _Version        =   655364
            _ExtentX        =   11245
            _ExtentY        =   503
            _StockProps     =   14
            Caption         =   "CONDITION"
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientColorLight=   16711680
            GradientColorDark=   16711680
            ForeColor       =   -2147483630
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2685
         Left            =   780
         ScaleHeight     =   2655
         ScaleWidth      =   5385
         TabIndex        =   12
         Top             =   5640
         Width           =   5415
         Begin VB.TextBox txtSPE 
            Height          =   900
            Left            =   90
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   1650
            Width           =   5205
         End
         Begin VB.TextBox txtKMS 
            Height          =   360
            Left            =   3660
            TabIndex        =   19
            Top             =   330
            Width           =   945
         End
         Begin VB.TextBox txtEVERY 
            Height          =   360
            Left            =   690
            TabIndex        =   18
            Top             =   330
            Width           =   1425
         End
         Begin VB.PictureBox Picture8 
            BorderStyle     =   0  'None
            Height          =   645
            Left            =   90
            ScaleHeight     =   645
            ScaleWidth      =   3615
            TabIndex        =   13
            Top             =   720
            Width           =   3615
            Begin VB.OptionButton optDEL 
               Caption         =   "OTHERS"
               Height          =   255
               Index           =   3
               Left            =   0
               MouseIcon       =   "Accmaintenance.frx":4E8A
               MousePointer    =   99  'Custom
               TabIndex        =   17
               Top             =   390
               Width           =   1305
            End
            Begin VB.OptionButton optDEL 
               Caption         =   "GAS STATION"
               Height          =   255
               Index           =   1
               Left            =   2100
               MouseIcon       =   "Accmaintenance.frx":4FDC
               MousePointer    =   99  'Custom
               TabIndex        =   16
               Top             =   30
               Width           =   1395
            End
            Begin VB.OptionButton optDEL 
               Caption         =   "3-STAR SHOP"
               Height          =   255
               Index           =   2
               Left            =   2070
               MouseIcon       =   "Accmaintenance.frx":512E
               MousePointer    =   99  'Custom
               TabIndex        =   15
               Top             =   390
               Width           =   1665
            End
            Begin VB.OptionButton optDEL 
               Caption         =   "DEALER"
               Height          =   255
               Index           =   0
               Left            =   0
               MouseIcon       =   "Accmaintenance.frx":5280
               MousePointer    =   99  'Custom
               TabIndex        =   14
               Top             =   30
               Value           =   -1  'True
               Width           =   1275
            End
         End
         Begin VB.Label lblcap 
            Caption         =   "OTHERS (PLS. SPECIFY)"
            Height          =   240
            Index           =   38
            Left            =   120
            TabIndex        =   25
            Top             =   1410
            Width           =   1710
         End
         Begin VB.Label lblcap 
            Caption         =   "KMS"
            Height          =   240
            Index           =   41
            Left            =   4710
            TabIndex        =   24
            Top             =   375
            Width           =   285
         End
         Begin VB.Label lblcap 
            Caption         =   "MONTH/S, EVERY "
            Height          =   240
            Index           =   39
            Left            =   2310
            TabIndex        =   23
            Top             =   345
            Width           =   1335
         End
         Begin VB.Label lblcap 
            Caption         =   "EVERY"
            Height          =   240
            Index           =   25
            Left            =   90
            TabIndex        =   22
            Top             =   345
            Width           =   480
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   285
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   6345
            _Version        =   655364
            _ExtentX        =   11192
            _ExtentY        =   503
            _StockProps     =   14
            Caption         =   "VEHICLE MAINTENANCE"
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientColorLight=   16711680
            GradientColorDark=   16711680
            ForeColor       =   -2147483630
         End
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   8370
         MouseIcon       =   "Accmaintenance.frx":53D2
         MousePointer    =   99  'Custom
         Picture         =   "Accmaintenance.frx":5524
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Save Entry"
         Top             =   6030
         Width           =   645
      End
      Begin VB.CommandButton cmdCloseTerm 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   9480
         MouseIcon       =   "Accmaintenance.frx":5874
         MousePointer    =   99  'Custom
         Picture         =   "Accmaintenance.frx":59C6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel Entry"
         Top             =   6750
         Width           =   645
      End
      Begin VB.Label lblcap 
         Caption         =   "VEHICLE SPEED"
         Height          =   240
         Index           =   37
         Left            =   180
         TabIndex        =   92
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label lblcap 
         Caption         =   "ALL"
         Height          =   240
         Index           =   42
         Left            =   1440
         TabIndex        =   91
         Top             =   360
         Width           =   270
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
      Left            =   2370
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   6
      Top             =   2460
      Width           =   1800
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   720
         MouseIcon       =   "Accmaintenance.frx":5B18
         MousePointer    =   99  'Custom
         Picture         =   "Accmaintenance.frx":5C6A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   30
         MouseIcon       =   "Accmaintenance.frx":5FA8
         MousePointer    =   99  'Custom
         Picture         =   "Accmaintenance.frx":60FA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Save New Password"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.TextBox txtConfirm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   150
      MaxLength       =   22
      PasswordChar    =   "l"
      TabIndex        =   5
      Top             =   2010
      Width           =   3645
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   150
      MaxLength       =   22
      PasswordChar    =   "l"
      TabIndex        =   3
      Top             =   1380
      Width           =   3645
   End
   Begin VB.TextBox txtCurrentpass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   150
      MaxLength       =   22
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   720
      Width           =   3645
   End
   Begin VB.Label Label3 
      Caption         =   "CHANGE THE  PASWORD OF YOUR ACCOUNT:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   120
      TabIndex        =   93
      Top             =   120
      Width           =   3630
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Type the new password again to confirm"
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
      Height          =   225
      Left            =   150
      TabIndex        =   4
      Top             =   1770
      Width           =   3315
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Type your Current Password"
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
      Left            =   150
      TabIndex        =   0
      Top             =   450
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Type a new password"
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
      Left            =   150
      TabIndex        =   2
      Top             =   1110
      Width           =   3375
   End
End
Attribute VB_Name = "frmAccMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode:

    If LTrim(RTrim(txtCurrentpass)) = "" Then
        ShowIsRequiredMsg " Current Password"
        On Error Resume Next
        txtCurrentpass.SetFocus
        Exit Sub
    End If
    
    If LTrim(RTrim(txtPassword)) = "" Then
        ShowIsRequiredMsg "New Password"
        On Error Resume Next
        txtPassword.SetFocus
        Exit Sub
    End If
    
    If LTrim(RTrim(txtConfirm)) = "" Then
        ShowIsRequiredMsg "Confirm Password"
        On Error Resume Next
        txtConfirm.SetFocus
        Exit Sub
    End If

    If LOGPASS <> txtCurrentpass Then
        MessagePop RecSaveError, "Password Not Match ", "Password doesn't match to your current Password"
        On Error Resume Next
        txtCurrentpass.SetFocus
        Exit Sub
    End If

    If txtConfirm <> txtPassword Then
        MessagePop RecSaveError, "Password Not Match ", "New Password doesn't match with your confirmed Password."
        On Error Resume Next
        txtConfirm.SetFocus
        Exit Sub
    End If
    
    With wizVar
        SQL_STATEMENT = "UPDATE ALL_RAMS_USERS SET PASSWORD='" & .EncryptAccess(txtPassword) & "' WHERE USERID=" & LOGID
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "PASSWORD MAINTENANCE", SQL_STATEMENT, "", "", "", "", ""
        ShowSuccessFullyUpdated
        cmdCancel.Value = True
    End With
    Exit Sub
    
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub txtCurrentpass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
End Sub

