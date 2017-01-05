VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Trans_MRR1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK TRANSFER OUT"
   ClientHeight    =   6450
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   11535
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
   Icon            =   "StockTransfer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   11535
   Begin VB.PictureBox picBottoms 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   11535
      TabIndex        =   3
      Top             =   5400
      Width           =   11535
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   960
         Top             =   150
      End
      Begin Crystal.CrystalReport rptMRR 
         Left            =   3300
         Top             =   330
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
         Left            =   9810
         ScaleHeight     =   885
         ScaleWidth      =   1800
         TabIndex        =   4
         Top             =   60
         Width           =   1800
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            CausesValidation=   0   'False
            Height          =   795
            Left            =   930
            MouseIcon       =   "StockTransfer.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Cancel"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   240
            MouseIcon       =   "StockTransfer.frx":0D5A
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":0EAC
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Save this Record"
            Top             =   30
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
         Left            =   3270
         ScaleHeight     =   945
         ScaleWidth      =   8805
         TabIndex        =   7
         Top             =   30
         Width           =   8805
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   7380
            MouseIcon       =   "StockTransfer.frx":11FC
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":134E
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Exit Window"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   795
            Left            =   6690
            MouseIcon       =   "StockTransfer.frx":16B4
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":1806
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Print this Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   300
            MouseIcon       =   "StockTransfer.frx":1B6C
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":1CBE
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Delete Selected Record"
            Top             =   60
            Visible         =   0   'False
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
            Left            =   6000
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "StockTransfer.frx":1FE9
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":213B
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Cancel this Transaction"
            Top             =   60
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
            Left            =   5310
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "StockTransfer.frx":2475
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":25C7
            Style           =   1  'Graphical
            TabIndex        =   102
            ToolTipText     =   "Post this Transaction"
            Top             =   60
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
            Left            =   4620
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "StockTransfer.frx":28EC
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":2A3E
            Style           =   1  'Graphical
            TabIndex        =   104
            ToolTipText     =   "Unpost this Transaction"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   3930
            MouseIcon       =   "StockTransfer.frx":2D83
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":2ED5
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Edit Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   3240
            MouseIcon       =   "StockTransfer.frx":3231
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":3383
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Add Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   2550
            MouseIcon       =   "StockTransfer.frx":3696
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":37E8
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Find a Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1860
            MouseIcon       =   "StockTransfer.frx":3AE2
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":3C34
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Move to Next Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   1170
            MouseIcon       =   "StockTransfer.frx":3F8C
            MousePointer    =   99  'Custom
            Picture         =   "StockTransfer.frx":40DE
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Move to Previous Record"
            Top             =   60
            Width           =   705
         End
      End
   End
   Begin VB.PictureBox picTops 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   11535
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton Command4 
         Caption         =   "::"
         Height          =   345
         Left            =   11040
         TabIndex        =   108
         ToolTipText     =   "Edit Transaction Date"
         Top             =   300
         Width           =   345
      End
      Begin MSComCtl2.DTPicker txtDateTransfered 
         Height          =   345
         Left            =   9120
         TabIndex        =   100
         Top             =   300
         Width           =   1935
         _ExtentX        =   3413
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
         CalendarTitleBackColor=   4194304
         CalendarTitleForeColor=   16777215
         Format          =   50790401
         CurrentDate     =   39258
      End
      Begin VB.TextBox txtSDNO 
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
         Left            =   7140
         MaxLength       =   6
         TabIndex        =   97
         Top             =   300
         Width           =   1935
      End
      Begin VB.ComboBox cboEntityFrom 
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
         Left            =   12030
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   930
         Visible         =   0   'False
         Width           =   3405
      End
      Begin VB.ComboBox cboEntityTo 
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
         Left            =   150
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   300
         Width           =   3435
      End
      Begin VB.Label LABALLOWREPRINT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   2190
         TabIndex        =   109
         Top             =   270
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label labStatus 
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
         Height          =   510
         Left            =   5040
         TabIndex        =   105
         Top             =   180
         Width           =   1965
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   9090
         TabIndex        =   101
         Top             =   60
         Width           =   390
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "STR SD No"
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
         Left            =   7110
         TabIndex        =   98
         Top             =   30
         Width           =   900
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         TabIndex        =   16
         Top             =   60
         Width           =   210
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   990
         Visible         =   0   'False
         Width           =   1140
      End
   End
   Begin VB.PictureBox picVehicleReceving 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4665
      Left            =   0
      ScaleHeight     =   4665
      ScaleWidth      =   11535
      TabIndex        =   29
      Top             =   735
      Width           =   11535
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   60
         TabIndex        =   30
         Top             =   -60
         Width           =   11415
         Begin VB.CommandButton cmdSelectVehicles 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Select Vehicles"
            Height          =   375
            Left            =   8640
            MaskColor       =   &H00400000&
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Select Vehicles"
            Top             =   420
            Width           =   2595
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1995
            Left            =   60
            ScaleHeight     =   1995
            ScaleWidth      =   8535
            TabIndex        =   79
            Top             =   120
            Width           =   8535
            Begin VB.TextBox txtV_Color 
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
               Left            =   5400
               TabIndex        =   96
               Top             =   1260
               Width           =   2925
            End
            Begin VB.TextBox txtV_ProdNo 
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
               Left            =   5400
               TabIndex        =   94
               Top             =   900
               Width           =   2925
            End
            Begin VB.TextBox txtV_EngineNo 
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
               Left            =   5400
               TabIndex        =   91
               Top             =   1620
               Width           =   2925
            End
            Begin VB.TextBox txtV_ModelDescript 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1245
               TabIndex        =   85
               Top             =   90
               Width           =   7125
            End
            Begin VB.TextBox txtV_IgnKey 
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
               Left            =   1260
               TabIndex        =   84
               Top             =   1260
               Width           =   2565
            End
            Begin VB.TextBox txtV_SerialNo 
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
               Left            =   1260
               TabIndex        =   83
               Top             =   1620
               Width           =   2565
            End
            Begin VB.TextBox txtV_VINo 
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
               Left            =   1260
               TabIndex        =   82
               Top             =   900
               Width           =   2565
            End
            Begin VB.TextBox txtV_Make 
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
               Left            =   1245
               TabIndex        =   81
               Top             =   480
               Width           =   2535
            End
            Begin VB.TextBox txtV_Model 
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
               Left            =   3780
               TabIndex        =   80
               Top             =   480
               Width           =   1830
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
               Left            =   4920
               TabIndex        =   95
               Top             =   1320
               Width           =   450
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Vehicle Stock No"
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
               Left            =   3900
               TabIndex        =   93
               Top             =   960
               Width           =   1440
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
               Height          =   225
               Left            =   4500
               TabIndex        =   92
               Top             =   1680
               Width           =   840
            End
            Begin VB.Label Label9 
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
               Left            =   420
               TabIndex        =   90
               Top             =   1680
               Width           =   765
            End
            Begin VB.Label Label5 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "CS No"
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
               Left            =   660
               TabIndex        =   89
               Top             =   1320
               Width           =   510
            End
            Begin VB.Label Label10 
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
               Left            =   720
               TabIndex        =   88
               Top             =   960
               Width           =   435
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Make/Model"
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
               TabIndex        =   87
               Top             =   540
               Width           =   1020
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
               Left            =   180
               TabIndex        =   86
               Top             =   120
               Width           =   975
            End
         End
         Begin VB.TextBox txtV_Battery 
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
            Left            =   8640
            TabIndex        =   33
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtV_TireSize 
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
            Left            =   8640
            TabIndex        =   32
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label Label7 
            Caption         =   "Select Vehicles From The List"
            Height          =   195
            Left            =   8700
            TabIndex        =   99
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Battery"
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
            Left            =   8640
            TabIndex        =   35
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Tire Size"
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
            Left            =   8640
            TabIndex        =   34
            Top             =   1440
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Other Items Transferred"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   2475
         Left            =   60
         TabIndex        =   36
         Top             =   2100
         Width           =   3615
         Begin VB.TextBox txt_OI_14 
            Height          =   315
            Left            =   2040
            TabIndex        =   77
            Top             =   2115
            Width           =   1455
         End
         Begin VB.TextBox txt_OI_13 
            Height          =   315
            Left            =   2040
            TabIndex        =   64
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txt_OI_12 
            Height          =   315
            Left            =   2040
            TabIndex        =   63
            Top             =   1476
            Width           =   1455
         End
         Begin VB.TextBox txt_OI_11 
            Height          =   315
            Left            =   2040
            TabIndex        =   62
            Top             =   1152
            Width           =   1455
         End
         Begin VB.TextBox txt_OI_10 
            Height          =   315
            Left            =   2040
            TabIndex        =   61
            Top             =   828
            Width           =   1455
         End
         Begin VB.TextBox txt_OI_9 
            Height          =   315
            Left            =   2040
            TabIndex        =   60
            Top             =   504
            Width           =   1455
         End
         Begin VB.TextBox txt_OI_8 
            Height          =   315
            Left            =   2040
            TabIndex        =   59
            Top             =   180
            Width           =   1455
         End
         Begin VB.CheckBox chk_OI_7 
            Caption         =   "Service Manual"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   2145
            Width           =   1575
         End
         Begin VB.CheckBox chk_OI_6 
            Caption         =   "Owner's Manual"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   1830
            Width           =   2115
         End
         Begin VB.CheckBox chk_OI_5 
            Caption         =   "Warranty Manual"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   1506
            Width           =   1875
         End
         Begin VB.CheckBox chk_OI_4 
            Caption         =   "Set of Standard Tools"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   1182
            Width           =   1935
         End
         Begin VB.CheckBox chk_OI_3 
            Caption         =   "Cigar Lighter"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   858
            Width           =   1275
         End
         Begin VB.CheckBox chk_OI_2 
            Caption         =   "Keys"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   534
            Width           =   1155
         End
         Begin VB.CheckBox chk_OI_1 
            Caption         =   "Spare Tire"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   210
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
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
         ForeColor       =   &H00000040&
         Height          =   2475
         Left            =   3720
         TabIndex        =   65
         Top             =   2100
         Width           =   3795
         Begin VB.TextBox txt_AC_12 
            Height          =   315
            Left            =   2220
            TabIndex        =   78
            Top             =   1950
            Width           =   1455
         End
         Begin VB.CheckBox chk_AC_1 
            Caption         =   "Air Con"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   210
            Width           =   1695
         End
         Begin VB.CheckBox chk_AC_2 
            Caption         =   "AirCon Warranty Manual"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   577
            Width           =   2070
         End
         Begin VB.CheckBox chk_AC_3 
            Caption         =   "Stereo"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   944
            Width           =   1275
         End
         Begin VB.CheckBox chk_AC_4 
            Caption         =   "Antennae"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   1311
            Width           =   2055
         End
         Begin VB.CheckBox chk_AC_5 
            Caption         =   "Speaker"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   1620
            Width           =   1875
         End
         Begin VB.CheckBox chk_AC_6 
            Caption         =   "Stero Manual"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1980
            Width           =   2115
         End
         Begin VB.TextBox txt_AC_7 
            Height          =   315
            Left            =   2220
            TabIndex        =   70
            Top             =   180
            Width           =   1455
         End
         Begin VB.TextBox txt_AC_8 
            Height          =   315
            Left            =   2220
            TabIndex        =   69
            Top             =   534
            Width           =   1455
         End
         Begin VB.TextBox txt_AC_9 
            Height          =   315
            Left            =   2220
            TabIndex        =   68
            Top             =   888
            Width           =   1455
         End
         Begin VB.TextBox txt_AC_10 
            Height          =   315
            Left            =   2220
            TabIndex        =   67
            Top             =   1242
            Width           =   1455
         End
         Begin VB.TextBox txt_AC_11 
            Height          =   315
            Left            =   2220
            TabIndex        =   66
            Top             =   1596
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   2445
         Left            =   7560
         TabIndex        =   37
         Top             =   2100
         Width           =   3915
         Begin VB.TextBox txtRemarks 
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
            Height          =   2130
            Left            =   60
            MaxLength       =   600
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   210
            Width           =   3795
         End
      End
      Begin VB.Frame fraPrintingDetails 
         Caption         =   "Signatories"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   4620
         Left            =   11580
         TabIndex        =   39
         Top             =   2100
         Width           =   3900
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cancel"
            Height          =   315
            Left            =   3000
            MaskColor       =   &H00400000&
            Style           =   1  'Graphical
            TabIndex        =   107
            ToolTipText     =   "Cancel"
            Top             =   240
            Width           =   795
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Defaults"
            Height          =   315
            Left            =   2190
            MaskColor       =   &H00400000&
            Style           =   1  'Graphical
            TabIndex        =   106
            ToolTipText     =   "Defaults"
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txtSIG_PreparedBy 
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
            Left            =   120
            MaxLength       =   70
            TabIndex        =   45
            Top             =   720
            Width           =   3675
         End
         Begin VB.TextBox txtSIG_CheckedBy 
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
            Left            =   120
            MaxLength       =   70
            TabIndex        =   44
            Top             =   1320
            Width           =   3675
         End
         Begin VB.TextBox txtSIG_ApprovedBy 
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
            Left            =   120
            MaxLength       =   70
            TabIndex        =   43
            Top             =   1920
            Width           =   3675
         End
         Begin VB.TextBox txtSIG_DeliveredBy 
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
            Left            =   120
            MaxLength       =   70
            TabIndex        =   42
            Top             =   2520
            Width           =   3675
         End
         Begin VB.TextBox txtSIG_ReleasedBy 
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
            Left            =   120
            MaxLength       =   70
            TabIndex        =   41
            Top             =   3120
            Width           =   3675
         End
         Begin VB.TextBox txtSIG_PostedBy 
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
            Left            =   120
            MaxLength       =   70
            TabIndex        =   40
            Top             =   3720
            Width           =   3675
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
            Left            =   120
            TabIndex        =   51
            Top             =   480
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
            Left            =   120
            TabIndex        =   50
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Approved By"
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
            TabIndex        =   49
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Delivered By"
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
            TabIndex        =   48
            Top             =   2280
            Width           =   1050
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Released By"
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
            TabIndex        =   47
            Top             =   2880
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Posted By"
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
            TabIndex        =   46
            Top             =   3480
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox picViewVehicles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4860
      Left            =   1200
      ScaleHeight     =   4830
      ScaleWidth      =   9720
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   9750
      Begin XtremeReportControl.ReportControl lvViewVehicles 
         Height          =   3405
         Left            =   60
         TabIndex        =   20
         Top             =   750
         Width           =   9540
         _Version        =   655364
         _ExtentX        =   16828
         _ExtentY        =   6006
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "&Cancel"
         Height          =   600
         Index           =   2
         Left            =   8850
         MouseIcon       =   "StockTransfer.frx":443D
         MousePointer    =   99  'Custom
         Picture         =   "StockTransfer.frx":458F
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Cancel"
         Top             =   4200
         Width           =   645
      End
      Begin VB.TextBox txtFilterViewVehicles 
         Height          =   375
         Left            =   5460
         TabIndex        =   22
         Top             =   375
         Width           =   4155
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
         TabIndex        =   21
         Top             =   15
         Width           =   285
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   600
         Left            =   8220
         MouseIcon       =   "StockTransfer.frx":48CD
         MousePointer    =   99  'Custom
         Picture         =   "StockTransfer.frx":4A1F
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Select"
         Top             =   4200
         Width           =   645
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "C#= Conduction Sticker No . P#= Production No. E# = Engine No ."
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   60
         TabIndex        =   26
         Top             =   4335
         Width           =   7515
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Left            =   0
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   450
         Width           =   2505
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "F# = Frame No . V#= VIN No .S#=Serial No"
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   75
         TabIndex        =   23
         Top             =   4560
         Width           =   7515
      End
   End
End
Attribute VB_Name = "frmSMIS_Trans_MRR1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsST                                                              As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim WithEvents SearchMaster                                           As frmSMIS_Mis_SearchMaster
Attribute SearchMaster.VB_VarHelpID = -1

Function DetectATMT(strx)
    Dim I                                                             As Integer
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

Sub SearchID(XXX)
    On Error GoTo Errorcode
    rsST.MoveFirst
    rsST.Find ("ID=" & XXX)
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Sub InitData()
    ReportControlAddColumnHeader lvViewVehicles, "MAKE,MODEL,YEAR,DESCRIPTION, C#, P#, E#,F#,V#,S#, COLOR, #MCODE"
    ReportControlPaintManager lvViewVehicles
    ResizeColumnHeader lvViewVehicles, "8,6,6,20,8,8,8,8,8,8,8,8"
End Sub

Sub InitMemVars()
    Dim cntrl                                                         As Control
    For Each cntrl In Me.ControlS
        If TypeOf cntrl Is CheckBox Then
            cntrl.Value = 0
        ElseIf TypeOf cntrl Is TextBox Then
            cntrl.Text = ""
        ElseIf TypeOf cntrl Is ComboBox Then
            cntrl.Text = ""
        End If
    Next
    LABALLOWREPRINT = ""
    labStatus.Caption = ""
    txtDateTransfered.Enabled = True
    txtDateTransfered = DateValue(LOGDATE)

    Combo_Loadval cboEntityTo, gconDMIS.Execute("select distinct entity_to from smis_stocktransfer where entity_to is not null")
End Sub

Sub LoadDefaultSignatories()
    Dim rsSig                                                         As ADODB.Recordset
    Set rsSig = gconDMIS.Execute("Select  TOP 1  * from SMIS_Signatories Where UsedIn='STOCK TRANSFER'")
    If Not rsSig.EOF Or Not rsSig.BOF Then
        txtSIG_ApprovedBy = Null2String(rsSig!GeneralManager)
        txtSIG_CheckedBy = Null2String(rsSig!CheckedBy)
        txtSIG_DeliveredBy = Null2String(rsSig!DeliveredBy)
        txtSIG_PostedBy = LOGNAME
        txtSIG_PreparedBy = Null2String(rsSig!PreparedBy)
        txtSIG_ReleasedBy = Null2String(rsSig!SalesDispatcher)
    End If
End Sub

Private Sub Command4_Click()
    If Function_Access(LOGID, "ACESS_SYSTEM", "STOCK TRANSFER") = False Then Exit Sub
    txtDateTransfered.Enabled = True: txtDateTransfered.SetFocus
End Sub

Private Sub cboEntityTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then: KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "STOCK TRANSFER") = False Then Exit Sub
    On Error GoTo Errorcode:

    If MsgBox("Do you Want to Cancel this Transaction ", vbYesNo + vbExclamation, "Confirm Posting") = vbNo Then Exit Sub
    cmdCancelCO.Enabled = False

    'gconDMIS.Execute ("UPDate SMIS_StockTransfer Set Status='C'  Where ID=" & labid)
    SQL_STATEMENT = ("UPDate SMIS_StockTransfer Set Status='C'  Where ID=" & labid)

    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "C", "STOCK TRANSFER", SQL_STATEMENT, Null2String(labid), "", "CS No:" & txtV_IgnKey, "", ""

    gconDMIS.Execute ("Update SMIS_MRRINV SET ISTATUS='O' , datereleased =NULL, RELEASED=0  WHERE ignkey=" & N2Str2Null(rsST!CSNO))
    LogAudit "C", "VEHICLE TRANSFER", cboEntityTo & " MODEL: " & txtV_ModelDescript & " CS NO:" & txtV_IgnKey
    rsST.Requery
    rsST.Find ("ID=" & labid)
    StoreMemVars
    MessagePop RecSaveOk, "Transaction Cancelled", "Transaction Sucessfully Cancelled"





    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdFind_Click()
    frmSMIS_SearchVehicleStockTransfer.SearchTab = 1
    frmSMIS_SearchVehicleStockTransfer.Show 1



End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "STOCK TRANSFER") = False Then Exit Sub
    On Error GoTo Errorcode:
    If MsgBox("Do you Want to Post this Transaction ", vbYesNo + vbExclamation, "Confirm Posting") = vbNo Then Exit Sub
    cmdCancelCO.Enabled = True


    'gconDMIS.Execute ("UPDate SMIS_StockTransfer  Set Status='P' Where ID=" & labid)

    SQL_STATEMENT = ("UPDate SMIS_StockTransfer  Set Status='P' Where ID=" & labid)

    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "P", "STOCK TRANSFER", SQL_STATEMENT, Null2String(labid), "", "CS No:" & txtV_IgnKey, "", ""

    LogAudit "P", "VEHICLE TRANSFER", cboEntityTo & " MODEL: " & txtV_ModelDescript & " CS NO:" & txtV_IgnKey
    rsST.Requery
    rsST.Find ("ID=" & labid)
    StoreMemVars
    MessagePop RecSaveOk, "Transaction Posted", "Transaction Sucessfully Posted"
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdUnPost_Click()
    Dim rsCheck As ADODB.Recordset
    
    If Function_Access(LOGID, "Acess_UnPost", "STOCK TRANSFER") = False Then Exit Sub
    On Error GoTo Errorcode:

    'UPDATED BY: JUN
    'DATE UPDATED: 10-23-2008
    'DESCRIPTION: USER WILL NOT ALLOW TO UNPOST TRANSACTION IF IT IS ALREADY POSTED IN STOCK TRANSFER IN
    Set rsCheck = gconDMIS.Execute("Select * from SMIS_Stocktransfer where CSNO = '" & txtV_IgnKey & "' and ENTITY_FROM IS NOT NULL and STATUS = 'P'")
    If Not rsCheck.EOF And Not rsCheck.BOF Then
        MsgBox "You cannot Unpost this Unit Already" & vbCrLf & "Posted in STOCK TRANSFER IN", vbInformation, "INFORMATION"
        Exit Sub
    Else
        If MsgBox("Do you Want to Un Post this Transaction ", vbYesNo + vbExclamation, "Confirm Un-Posting") = vbNo Then Exit Sub
        cmdCancelCO.Enabled = False
        
        SQL_STATEMENT = ("UPDATE SMIS_STOCKTRANSFER SET STATUS='U' WHERE ID=" & labid)
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "U", "STOCK TRANSFER", SQL_STATEMENT, Null2String(labid), "", "CS No:" & txtV_IgnKey, "", ""
        LogAudit "U", "VEHICLE TRANSFER", cboEntityTo & " MODEL: " & txtV_ModelDescript & " CS NO:" & txtV_IgnKey
    End If
    rsST.Requery
    rsST.Find ("ID=" & labid)
    StoreMemVars
    MessagePop RecSaveOk, "TRANSACTION POSTED", "TRANSACTION SUCESSFULLY UNPOSTED"
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub Command1_Click()
    LoadDefaultSignatories
End Sub

Private Sub Command2_Click()
    StoreMemVars
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (STOCK TRANSFER)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "STOCK TRANSFER")
            'End If
    End Select

End Sub

Private Sub lvViewVehicles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSelect_Click
    End If
End Sub

Private Sub lvViewVehicles_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdSelect_Click
End Sub

Private Sub txtFilterViewVehicles_Change()
    lvViewVehicles.FilterText = txtFilterViewVehicles
    lvViewVehicles.Populate
End Sub

Private Sub txtV_ModelDescript_Change()
    'UDPATING CODE    : AXP-652005106
    'UDPATING CODE      :   AXP-672007543
    If AddorEdit = "" Then: Exit Sub
    If RTrim(LTrim(txtV_ModelDescript)) = "" Then: Exit Sub
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim rsModelCode                                                   As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("Select MODEL from ALL_MODEL where descript=" & N2Str2Null(txtV_ModelDescript))
    If Not (TEMPRS.BOF Or TEMPRS.EOF) Then
        txtV_Model = Null2String(TEMPRS!Model)
        Set rsModelCode = gconDMIS.Execute("select CODE FROM ALL_ModelCode where description=" & N2Str2Null(txtV_Model))
        If Not rsModelCode.EOF Or Not rsModelCode.BOF Then
        End If
    End If
    Set TEMPRS = Nothing
    Set rsModelCode = Nothing
End Sub

Private Sub txtV_ModelDescript_Click()
    txtV_ModelDescript_Change
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "STOCK TRANSFER") = False Then Exit Sub
    AddorEdit = "ADD"
    picVehicleReceving.Enabled = True
    picTops.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    InitMemVars
    txtSDNO = GenerateCode("SMIS_StockTransfer", "SDNO", "000000")
End Sub

Private Sub cmdCancel_Click()
    txtV_IgnKey.Enabled = True
    txtV_ProdNo.Enabled = True
    txtV_SerialNo.Enabled = True
    'UPDATED BY: JUN------------------------------
    'DATE UPDATED: 11-20-2008
    cmdSelectVehicles.Enabled = True
    '---------------------------------------------
    AddorEdit = ""
    picTops.Enabled = False: picAdds.Visible = True: picSaves.Visible = False: picVehicleReceving.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdCancelViewVehicles_Click(Index As Integer)
    ShowHidePictureBox2 picViewVehicles, False
End Sub

''END CLOSE

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "STOCK TRANSFER") = False Then Exit Sub
    On Error GoTo Errorcode:

    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from SMIS_StockTransfer where id = " & labid.Caption


        gconDMIS.Execute ("Update SMIS_MRRINV SET ISTATUS='O' WHERE PRODNO=" & N2Str2Null(rsST!vSNO))

        ShowDeletedMsg
        rsRefresh
        StoreMemVars
    End If





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()

    If Function_Access(LOGID, "Acess_EDIT", "STOCK TRANSFER") = False Then Exit Sub
    On Error GoTo Errorcode:
    AddorEdit = "EDIT"
    picTops.Enabled = True
    picVehicleReceving.Enabled = True
    picSaves.Visible = True
    txtDateTransfered.Enabled = False
    picAdds.Visible = False
    'UPDATED BY: JUN------------------------------
    'DATE UPDATED: 11-20-2008
    'DESCRIPTION: IN ORDER NOT TO BE UPDATE OTHER VEHICLES IN STOCK
    cmdSelectVehicles.Enabled = False
    '---------------------------------------------
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    On Error GoTo Errorcode:

    rsST.MoveNext
    If rsST.EOF Then
        rsST.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars





    Exit Sub
Errorcode:
    ShowVBError
End Sub

''LISTIVEWS


Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:

    rsST.MovePrevious
    If rsST.BOF Then
        rsST.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "STOCK TRANSFER") = False Then Exit Sub

    If LABALLOWREPRINT <> "" Then
        If AllowReprint("STOCK TRANSFER") = False Then Exit Sub
    End If
    Screen.MousePointer = 11
    If COMPANY_CODE = "HBK" Then
        LoadSignatories "STOCK TRANSFER"


        rptMRR.Formulas(0) = "PREPARED_BY='" & (PreparedBy & "'")
        rptMRR.Formulas(1) = "DELIVERED_BY='" & DeliveredBy & "'"
        rptMRR.Formulas(2) = "CHECKED_BY='" & CheckedBy & "'"
        rptMRR.Formulas(3) = "RELEASED_BY='" & SalesDispatcher & "'"
        rptMRR.Formulas(4) = "APPROVED_BY='" & GeneralManager & "'"
        rptMRR.Formulas(5) = "POSTED_BY='" & FinancingManager & "'"
    End If
    'upadated by:   IEBV 07142010_0625pm
    'description:   For the HGC Report
    If (COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGO" Or COMPANY_CODE = "HAI") Then
        rptMRR.Formulas(0) = "CompanyName='" & COMPANY_NAME & "'"
        rptMRR.Formulas(1) = "CompanyAddress='" & COMPANY_ADDRESS & "'"
    End If
    '-----------------------------------------------------------------------
    PrintSQLReport rptMRR, SMIS_REPORT_PATH & "str.rpt", "{SST.ID} = " & labid, DMIS_REPORT_Connection, 1

    NEW_LogAudit "V", "STOCK TRANSFER", "", Null2String(labid), "", "CS No:" & txtV_IgnKey, "", ""


    LogAudit "V", "VEHICLE TRANSFER", cboEntityTo & " MODEL: " & txtV_ModelDescript & " CS NO:" & txtV_IgnKey
    Screen.MousePointer = 0
    gconDMIS.Execute ("update SMIS_STOCKTRANSFER SET PRINTED=1 WHERE ID=" & labid)
    rsRefresh
    rsST.Find ("id=" & labid)
    StoreMemVars
End Sub

Private Sub cmdSave_Click()

    On Error GoTo Errorcode:

    If RTrim(LTrim(txtSDNO)) = "" Then
        ShowIsRequiredMsg " Standard SD No. "
        On Error Resume Next
        txtSDNO.SetFocus
        Exit Sub
    End If

    '   If RTrim(LTrim(cboEntityFrom)) = "" Then
    '       ShowIsRequiredMsg " From . "
    '       On Error Resume Next
    '        cboEntityFrom.SetFocus
    '      Exit Sub
    '   End If

    If RTrim(LTrim(cboEntityTo)) = "" Then
        ShowIsRequiredMsg " To "
        On Error Resume Next
        cboEntityTo.SetFocus
        Exit Sub
    End If

    If RTrim(LTrim(txtV_IgnKey)) = "" Then
        ShowIsRequiredMsg " Please Select Vehicle  Detail From The List. "
        On Error Resume Next
        cmdSelectVehicles.SetFocus
        Exit Sub
    End If



    Dim lngcount                                                      As Integer
    Dim vtxt_Entity_To, vtxt_Entity_From, vtxt_SDNO                   As String
    Dim vtxt_Descriptions, vtxt_Make, vtxt_Model                      As String
    Dim vtxt_VINO, vtxt_CSNO, vtxt_SERIALNO, vtxt_ENGINENO, vtxt_VSNO, vtxt_Color As String
    Dim vtxt_BATTERY, vtxt_TireSize
    Dim vtxt_OI_F1, vtxt_OI_F2, vtxt_OI_F3, vtxt_OI_F4, vtxt_OI_F5, vtxt_OI_F6 As String
    Dim vtxt_OI_F7, vtxt_OI_F8, vtxt_OI_F9, vtxt_OI_F10, vtxt_OI_F11, vtxt_OI_F12, vtxt_OI_F13, vtxt_OI_F14 As String
    Dim vtxt_AC_F1, vtxt_AC_F2, vtxt_AC_F3, vtxt_AC_F4, vtxt_AC_F5, vtxt_AC_F6 As String
    Dim vtxt_AC_F7, vtxt_AC_F8, vtxt_AC_F9, vtxt_AC_F10, vtxt_AC_F11, vtxt_AC_F12 As String
    Dim PreparedBy, CheckedBy, ApprovedBy, DeliveredBy, RepleasedBy, PostedBy As String
    Dim SQL                                                           As String

    vtxt_Entity_To = N2Str2Null(cboEntityTo)
    vtxt_Entity_From = N2Str2Null(cboEntityFrom)
    vtxt_SDNO = N2Str2Null(txtSDNO)
    vtxt_Descriptions = N2Str2Null(txtV_ModelDescript)
    vtxt_Make = N2Str2Null(txtV_Make)
    vtxt_Model = N2Str2Null(txtV_Model)
    vtxt_VINO = N2Str2Null(txtV_VINo)
    vtxt_CSNO = N2Str2Null(txtV_IgnKey)
    vtxt_SERIALNO = N2Str2Null(txtV_SerialNo)
    vtxt_ENGINENO = N2Str2Null(txtV_EngineNo)
    vtxt_VSNO = N2Str2Null(txtV_ProdNo)
    vtxt_Color = N2Str2Null(txtV_Color)
    vtxt_BATTERY = N2Str2Null(txtV_Battery)
    vtxt_TireSize = N2Str2Null(txtV_TireSize)

    vtxt_OI_F1 = chk_OI_1.Value
    vtxt_OI_F2 = chk_OI_2.Value
    vtxt_OI_F3 = chk_OI_3.Value
    vtxt_OI_F4 = chk_OI_4.Value
    vtxt_OI_F5 = chk_OI_5.Value
    vtxt_OI_F6 = chk_OI_6.Value
    vtxt_OI_F7 = chk_OI_7.Value
    vtxt_OI_F8 = N2Str2Null(txt_OI_8)
    vtxt_OI_F9 = N2Str2Null(txt_OI_9)
    vtxt_OI_F10 = N2Str2Null(txt_OI_10)
    vtxt_OI_F11 = N2Str2Null(txt_OI_11)
    vtxt_OI_F12 = N2Str2Null(txt_OI_12)
    vtxt_OI_F13 = N2Str2Null(txt_OI_13)
    vtxt_OI_F14 = N2Str2Null(txt_OI_14)

    vtxt_AC_F1 = chk_AC_1.Value
    vtxt_AC_F2 = chk_AC_2.Value
    vtxt_AC_F3 = chk_AC_3.Value
    vtxt_AC_F4 = chk_AC_4.Value
    vtxt_AC_F5 = chk_AC_5.Value
    vtxt_AC_F6 = chk_AC_6.Value
    vtxt_AC_F7 = N2Str2Null(txt_AC_7)
    vtxt_AC_F8 = N2Str2Null(txt_AC_8)
    vtxt_AC_F9 = N2Str2Null(txt_AC_9)
    vtxt_AC_F10 = N2Str2Null(txt_AC_10)
    vtxt_AC_F11 = N2Str2Null(txt_AC_11)
    vtxt_AC_F12 = N2Str2Null(txt_AC_12)

    PreparedBy = N2Str2Null(txtSIG_PreparedBy)
    CheckedBy = N2Str2Null(txtSIG_CheckedBy)
    ApprovedBy = N2Str2Null(txtSIG_ApprovedBy)
    DeliveredBy = N2Str2Null(txtSIG_DeliveredBy)
    RepleasedBy = N2Str2Null(txtSIG_ReleasedBy)
    PostedBy = N2Str2Null(txtSIG_PostedBy)


    'gconDMIS.Execute ("Update SMIS_MRRINV SET ISTATUS='O' WHERE ignkey=" & vtxt_CSNO)
    Dim rsHanapID                                                     As ADODB.Recordset

    Set rsHanapID = New ADODB.Recordset
    Dim vID                                                           As String
    If AddorEdit = "ADD" Then

        SQL = "INSERT INTO SMIS_StockTransfer"
        SQL = SQL & " ( Deyt, "
        SQL = SQL & " Entity_To, Entity_From, SDNO, "
        SQL = SQL & " Descriptions, Make, Model, VINO, CSNO, SERIALNO, ENGINENO, VSNO, COLOR, "
        SQL = SQL & " BATTERY, TireSize, "
        SQL = SQL & " OI_F1, OI_F2, OI_F3, OI_F4, OI_F5, OI_F6, OI_F7, "
        SQL = SQL & " OI_F8, OI_F9, OI_F10, OI_F11, OI_F12, OI_F13, OI_F14, "
        SQL = SQL & " AC_F1, AC_F2, AC_F3, AC_F4, AC_F5, AC_F6, "
        SQL = SQL & " AC_F7, AC_F8, AC_F9, AC_F10, AC_F11, AC_F12, "
        SQL = SQL & " Notes ,"
        SQL = SQL & " PreparedBy, CheckedBy, ApprovedBy, "
        SQL = SQL & " DeliveredBy, RepleasedBy, PostedBy "
        SQL = SQL & " ) VALUES ("
        SQL = SQL & N2Str2Null(txtDateTransfered) & ","
        SQL = SQL & vtxt_Entity_To & " ," & vtxt_Entity_From & " ," & vtxt_SDNO & " ,"
        SQL = SQL & vtxt_Descriptions & " ," & vtxt_Make & " ," & vtxt_Model & " ," & vtxt_VINO & " ," & vtxt_CSNO & " ," & vtxt_SERIALNO & " ," & vtxt_ENGINENO & " ," & vtxt_VSNO & " ," & vtxt_Color & ", "
        SQL = SQL & vtxt_BATTERY & " ," & vtxt_TireSize & " ,"
        SQL = SQL & vtxt_OI_F1 & " ," & vtxt_OI_F2 & " ," & vtxt_OI_F3 & " ," & vtxt_OI_F4 & " ," & vtxt_OI_F5 & " ," & vtxt_OI_F6 & " ," & vtxt_OI_F7 & " ,"
        SQL = SQL & vtxt_OI_F8 & " ," & vtxt_OI_F9 & " ," & vtxt_OI_F10 & " ," & vtxt_OI_F11 & " ," & vtxt_OI_F12 & " ," & vtxt_OI_F13 & " ," & vtxt_OI_F14 & " ,"
        SQL = SQL & vtxt_AC_F1 & " ," & vtxt_AC_F2 & " ," & vtxt_AC_F3 & " ," & vtxt_AC_F4 & " ," & vtxt_AC_F5 & " ," & vtxt_AC_F6 & " ,"
        SQL = SQL & vtxt_AC_F7 & " ," & vtxt_AC_F8 & " ," & vtxt_AC_F9 & " ," & vtxt_AC_F10 & " ," & vtxt_AC_F11 & " ," & vtxt_AC_F12 & " , "
        SQL = SQL & N2Str2Null(txtRemarks) & " ,"
        SQL = SQL & PreparedBy & " ," & CheckedBy & " ," & ApprovedBy & ", "
        SQL = SQL & DeliveredBy & "," & RepleasedBy & " ," & PostedBy
        SQL = SQL & " )"
        gconDMIS.Execute SQL

        SQL_STATEMENT = SQL

        '******UPDATED BY RDC

        '******NEW LOG AUDIT*********
        NEW_LogAudit "A", "STOCK TRANSFER", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtV_IgnKey), "CSNO", "SMIS_STOCKTRANSFER"), "", "CS No:" & txtV_IgnKey, "", ""
        '****************************

        LogAudit "A", "VEHICLE TRANSFER", cboEntityTo & " MODEL: " & txtV_ModelDescript & " CS NO:" & txtV_IgnKey

    Else
        SQL = " UPDATE SMIS_StockTransfer SET "
        SQL = SQL & " Entity_To =" & vtxt_Entity_To & ", "
        SQL = SQL & " Entity_From =" & vtxt_Entity_From & ", "
        SQL = SQL & " Descriptions =" & vtxt_Descriptions & ", "
        SQL = SQL & " Make =" & vtxt_Make & ", "
        SQL = SQL & " Model =" & vtxt_Model & ", "
        SQL = SQL & " VINO =" & vtxt_VINO & ", "
        SQL = SQL & " CSNO =" & vtxt_CSNO & ", "
        SQL = SQL & " SERIALNO =" & vtxt_SERIALNO & ", "
        SQL = SQL & " ENGINENO =" & vtxt_ENGINENO & ", "
        SQL = SQL & " VSNO =" & vtxt_VSNO & ", "
        SQL = SQL & " COLOR =" & vtxt_Color & ", "
        SQL = SQL & " BATTERY =" & vtxt_BATTERY & ", "
        SQL = SQL & " TireSize =" & vtxt_TireSize & ", "
        SQL = SQL & " OI_F1 =" & vtxt_OI_F1 & ", "
        SQL = SQL & " OI_F2 =" & vtxt_OI_F2 & ", "
        SQL = SQL & " OI_F3 =" & vtxt_OI_F3 & ", "
        SQL = SQL & " OI_F4 =" & vtxt_OI_F4 & ", "
        SQL = SQL & " OI_F5  =" & vtxt_OI_F5 & ", "
        SQL = SQL & " OI_F6 =" & vtxt_OI_F6 & ", "
        SQL = SQL & " OI_F7 =" & vtxt_OI_F7 & ", "
        SQL = SQL & " OI_F8 =" & vtxt_OI_F8 & ", "
        SQL = SQL & " OI_F9 =" & vtxt_OI_F9 & ", "
        SQL = SQL & " OI_F10 =" & vtxt_OI_F10 & ", "
        SQL = SQL & " OI_F11 =" & vtxt_OI_F11 & ", "
        SQL = SQL & " OI_F12 =" & vtxt_OI_F12 & ", "
        SQL = SQL & " OI_F13 =" & vtxt_OI_F13 & ", "
        SQL = SQL & " OI_F14 =" & vtxt_OI_F14 & ", "
        SQL = SQL & " AC_F1 =" & vtxt_AC_F1 & ", "
        SQL = SQL & " AC_F2 =" & vtxt_AC_F2 & ", "
        SQL = SQL & " AC_F3 =" & vtxt_AC_F3 & ", "
        SQL = SQL & " AC_F4 =" & vtxt_AC_F4 & ", "
        SQL = SQL & " AC_F5 =" & vtxt_AC_F5 & ", "
        SQL = SQL & " AC_F6 =" & vtxt_AC_F6 & ", "
        SQL = SQL & " AC_F7 =" & vtxt_AC_F7 & ", "
        SQL = SQL & " AC_F8 =" & vtxt_AC_F8 & ", "
        SQL = SQL & " AC_F9 =" & vtxt_AC_F9 & ", "
        SQL = SQL & " AC_F10 =" & vtxt_AC_F10 & ", "
        SQL = SQL & " AC_F11 =" & vtxt_AC_F11 & ", "
        SQL = SQL & " AC_F12 =" & vtxt_AC_F12 & ", "
        SQL = SQL & " SDNO =" & vtxt_SDNO & ", "
        SQL = SQL & " DEYT =" & N2Str2Null(txtDateTransfered) & ", "
        SQL = SQL & " NOTES=" & N2Str2Null(txtRemarks) & " , "
        SQL = SQL & " PreparedBy= " & PreparedBy & ", "
        SQL = SQL & " CheckedBy= " & CheckedBy & ", "
        SQL = SQL & " ApprovedBy=" & ApprovedBy & ", "
        SQL = SQL & " DeliveredBy =" & DeliveredBy & ", "
        SQL = SQL & " RepleasedBy = " & RepleasedBy & ", "
        SQL = SQL & " PostedBy =" & PostedBy
        SQL = SQL & " WHERE ID=" & labid

        gconDMIS.Execute SQL
        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "STOCK TRANSFER", SQL_STATEMENT, Null2String(labid), "", "CS No:" & txtV_IgnKey, "", ""
        LogAudit "E", "VEHICLE TRANSFER", cboEntityTo & " MODEL: " & txtV_ModelDescript & " CS NO:" & txtV_IgnKey
    End If

    gconDMIS.Execute ("Update SMIS_MRRINV SET ISTATUS='T', released=1, datereleased=" & N2Str2Null(txtDateTransfered) & " WHERE ignkey=" & vtxt_CSNO)
    rsST.Requery
    If AddorEdit = "EDIT" Then
        rsST.Find ("ID=" & labid)
    End If
    cmdCancel.Value = True
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdSelect_Click()

    On Error GoTo ADDER:
    With lvViewVehicles.SelectedRows(0)
        txtV_Make = .Record(0).Value & ""
        txtV_Model = .Record(1).Value & ""
        txtV_ModelDescript = .Record(3).Value & ""
        txtV_IgnKey = .Record(4).Value & ""
        txtV_ProdNo = .Record(5).Value & ""
        txtV_EngineNo = .Record(6).Value & ""
        txtV_VINo = .Record(8).Value & ""
        txtV_SerialNo = .Record(9).Value & ""
        txtV_Color = .Record(10).Value & ""
        ShowHidePictureBox2 picViewVehicles, False

    End With
    Exit Sub
ADDER:
    MessagePop InfoVoid, "Selection Required", "There are to Record to Select "
    Err.Clear
End Sub

Private Sub cmdSelectVehicles_Click()

    'Make, Model , Yeer, Descript, ignkey,prodno,engineno,FrameNo,Vino,SerialNo,color,id
    flex_FillReportView gconDMIS.Execute("SELECT Make, Model , Yeer, Descript, ignkey,prodno,engineno,FrameNo,Vino,SerialNo,color,ModelCode, id , Transmission from SMIS_MRRINV where iStatus ='O'"), lvViewVehicles
    ShowHidePictureBox2 picViewVehicles, True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        'If picAccessories.Visible = True Then
        '    cmdCancelAcc_Click
        'ElseIf picFree.Visible = True Then
        '    cmdCancelFree_Click
        'End If
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11

    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    Set SearchMaster = New frmSMIS_Mis_SearchMaster

    rsRefresh
    InitData

    InitMemVars
    AddorEdit = ""
    picTops.Enabled = False
    picVehicleReceving.Enabled = False
    picSaves.Visible = False
    picAdds.Visible = True
    StoreMemVars
    Screen.MousePointer = 0

End Sub

Private Sub rsRefresh()
    Set rsST = New ADODB.Recordset
    rsST.CursorLocation = adUseClient
    Call rsST.Open("SELECT * from SMIS_StockTransfer WHERE ENTITY_TO IS NOT NULL ORDER BY ID desc", gconDMIS, adOpenKeyset)
End Sub

Private Sub StoreMemVars()
    If Not rsST.EOF And Not rsST.BOF Then
        LABALLOWREPRINT = Null2String(rsST!PRINTED)
        labid.Caption = rsST!ID
        txtSDNO = Null2String(rsST!SDNO)
        cboEntityFrom = Null2String(rsST!Entity_From)
        cboEntityTo = Null2String(rsST!ENTITY_TO)
        txtV_ModelDescript = Null2String(rsST!DESCRIPTIONS)
        txtV_Make = Null2String(rsST!Make)
        txtV_Model = Null2String(rsST!Model)
        txtV_IgnKey = Null2String(rsST!CSNO)
        txtV_ProdNo = Null2String(rsST!vSNO)
        txtV_SerialNo = Null2String(rsST!SERIALNO)
        txtV_VINo = Null2String(rsST!VINO)
        txtV_EngineNo = Null2String(rsST!EngineNo)
        txtV_Battery = Null2String(rsST!BATTERY)
        txtV_TireSize = Null2String(rsST!TireSize)
        txtV_Color = Null2String(rsST!Color)
        txtRemarks = Null2String(rsST!Notes)

        chk_OI_1 = Null2String(rsST!OI_F1)
        chk_OI_2 = Null2String(rsST!OI_F1)
        chk_OI_3 = Null2String(rsST!OI_F1)
        chk_OI_4 = Null2String(rsST!OI_F1)
        chk_OI_5 = Null2String(rsST!OI_F1)
        chk_OI_6 = Null2String(rsST!OI_F1)
        chk_OI_7 = Null2String(rsST!OI_F1)

        txt_OI_8 = Null2String(rsST!OI_F8)
        txt_OI_9 = Null2String(rsST!OI_F9)
        txt_OI_10 = Null2String(rsST!OI_F10)
        txt_OI_11 = Null2String(rsST!OI_F11)
        txt_OI_12 = Null2String(rsST!OI_F12)
        txt_OI_13 = Null2String(rsST!OI_F13)
        txt_OI_14 = Null2String(rsST!OI_F14)

        chk_AC_1 = Null2String(rsST!AC_F1)
        chk_AC_2 = Null2String(rsST!AC_F2)
        chk_AC_3 = Null2String(rsST!AC_F3)
        chk_AC_4 = Null2String(rsST!AC_F4)
        chk_AC_5 = Null2String(rsST!AC_F5)
        chk_AC_6 = Null2String(rsST!AC_F6)
        txt_AC_7 = Null2String(rsST!AC_F7)
        txt_AC_8 = Null2String(rsST!AC_F8)
        txt_AC_9 = Null2String(rsST!AC_F9)
        txt_AC_10 = Null2String(rsST!AC_F10)
        txt_AC_11 = Null2String(rsST!AC_F11)
        txt_AC_12 = Null2String(rsST!AC_F12)
        If IsDate(rsST!deyt) = True Then
            txtDateTransfered = rsST!deyt
        End If
        'POST UNPOST CANCEL
        If Null2String(rsST!STATUS) = "C" Then
            cmdCancelCO.Enabled = False:
            cmdUnPost.Enabled = False
            cmdPost.Enabled = False
            labStatus = "***CANCELLED***"
            cmdEdit.Enabled = False
            cmdPrint.Enabled = False
            cmdDelete.Enabled = False
        ElseIf Null2String(rsST!STATUS) = "P" Then
            cmdCancelCO.Enabled = False
            cmdUnPost.Enabled = True
            cmdPost.Enabled = False
            cmdDelete.Enabled = False
            labStatus = "***POSTED ***"
            cmdEdit.Enabled = False
            cmdPrint.Enabled = True
        Else
            cmdCancelCO.Enabled = True
            cmdUnPost.Enabled = False
            cmdPost.Enabled = True
            labStatus = ""
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            cmdPrint.Enabled = False
        End If

        txtSIG_PreparedBy = Null2String(rsST!PreparedBy)
        txtSIG_CheckedBy = Null2String(rsST!CheckedBy)
        txtSIG_ApprovedBy = Null2String(rsST!ApprovedBy)
        txtSIG_DeliveredBy = Null2String(rsST!DeliveredBy)
        txtSIG_ReleasedBy = Null2String(rsST!RepleasedBy)
        txtSIG_PostedBy = Null2String(rsST!PostedBy)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub Timer2_Timer()
    If labStatus.Caption <> "" Then
        If labStatus.Visible = True Then
            labStatus.Visible = False
        Else
            labStatus.Visible = True
        End If
    End If
End Sub

Private Sub txtV_EngineNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtV_Battery_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtV_IgnKey_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtV_TireSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtV_ProdNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtV_SerialNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtV_VINo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

