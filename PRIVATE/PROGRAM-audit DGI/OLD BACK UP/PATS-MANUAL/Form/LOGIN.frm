VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmLOGIN 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8880
   ClientLeft      =   735
   ClientTop       =   750
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "LOGIN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   PaletteMode     =   1  'UseZOrder
   Picture         =   "LOGIN.frx":030A
   ScaleHeight     =   592
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Begin VB.PictureBox Picture3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   90
      ScaleHeight     =   3465
      ScaleWidth      =   6825
      TabIndex        =   28
      Top             =   5220
      Width           =   6885
      Begin MSACAL.Calendar Calendar1 
         Height          =   3525
         Left            =   -30
         TabIndex        =   29
         Top             =   0
         Width           =   6885
         _Version        =   524288
         _ExtentX        =   12144
         _ExtentY        =   6218
         _StockProps     =   1
         BackColor       =   14141109
         Year            =   2008
         Month           =   1
         Day             =   22
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   16711680
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   32896
         GridLinesColor  =   -2147483626
         ShowDateSelectors=   0   'False
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   0   'False
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   0   'False
         TitleFontColor  =   255
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer6 
      Interval        =   500
      Left            =   2820
      Top             =   60
   End
   Begin MSComctlLib.ListView Listview1 
      Height          =   3255
      Left            =   420
      TabIndex        =   21
      Top             =   5430
      Visible         =   0   'False
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   3175
      EndProperty
      Picture         =   "LOGIN.frx":15FC4C
   End
   Begin VB.PictureBox picOT 
      BackColor       =   &H00D7C6B5&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   4365
      TabIndex        =   22
      Top             =   1800
      Visible         =   0   'False
      Width           =   4425
      Begin VB.OptionButton OptInOut 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   4
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   90
         Width           =   225
      End
      Begin VB.OptionButton OptInOut 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   5
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   60
         Width           =   225
      End
      Begin VB.Label Label7 
         BackColor       =   &H00D7C6B5&
         Caption         =   "Out - OT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   2670
         TabIndex        =   26
         Top             =   0
         Width           =   1665
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D7C6B5&
         Caption         =   "In - OT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   470
         TabIndex        =   25
         Top             =   30
         Width           =   1665
      End
   End
   Begin VB.Timer Timer5 
      Interval        =   10
      Left            =   17055
      Top             =   3420
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   17820
      Top             =   3420
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   8220
      ScaleHeight     =   105
      ScaleWidth      =   0
      TabIndex        =   20
      Top             =   5970
      Width           =   0
      Begin VB.Image Image1 
         Height          =   585
         Index           =   9
         Left            =   3210
         Picture         =   "LOGIN.frx":16C6FD
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   585
         Index           =   8
         Left            =   2850
         Picture         =   "LOGIN.frx":16CF2B
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   585
         Index           =   7
         Left            =   2490
         Picture         =   "LOGIN.frx":16D759
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   585
         Index           =   6
         Left            =   2130
         Picture         =   "LOGIN.frx":16DF87
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   585
         Index           =   5
         Left            =   1770
         Picture         =   "LOGIN.frx":16E7B5
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   585
         Index           =   4
         Left            =   1410
         Picture         =   "LOGIN.frx":16EFE3
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   585
         Index           =   3
         Left            =   1050
         Picture         =   "LOGIN.frx":16F811
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   585
         Index           =   2
         Left            =   690
         Picture         =   "LOGIN.frx":17003F
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   585
         Index           =   1
         Left            =   330
         Picture         =   "LOGIN.frx":17086D
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   585
         Index           =   0
         Left            =   0
         Picture         =   "LOGIN.frx":17109B
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   11490
      Top             =   30
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   11040
      Top             =   30
   End
   Begin VB.TextBox TxtTime 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   7110
      TabIndex        =   14
      Top             =   6930
      Width           =   4755
   End
   Begin VB.TextBox TxtDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   7140
      TabIndex        =   13
      Top             =   5670
      Width           =   4755
   End
   Begin VB.OptionButton OptInOut 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   3
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1290
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.OptionButton OptInOut 
      BackColor       =   &H00FF0000&
      Caption         =   "In - PM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   2
      Left            =   9705
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.OptionButton OptInOut 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   1
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1290
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.OptionButton OptInOut 
      BackColor       =   &H00FF0000&
      Caption         =   "In - AM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   0
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton CmdViewCards 
      BackColor       =   &H00E0E0E0&
      Caption         =   "View Cards"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   150
      Picture         =   "LOGIN.frx":1718C9
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7620
      Width           =   1215
   End
   Begin VB.CommandButton CmdMaintenance 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Maintenance"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   5550
      Picture         =   "LOGIN.frx":171BD3
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7620
      Width           =   1215
   End
   Begin VB.CommandButton CmdAbsent 
      BackColor       =   &H00E0E0E0&
      Caption         =   "View Absentees"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   2850
      Picture         =   "LOGIN.frx":171EDD
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7620
      Width           =   1215
   End
   Begin VB.CommandButton CmdEditCards 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Edit Cards"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   1500
      Picture         =   "LOGIN.frx":1721E7
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7620
      Width           =   1215
   End
   Begin VB.CommandButton CmdSummary 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Attendance Summary"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   4200
      Picture         =   "LOGIN.frx":1724F1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7620
      Width           =   1215
   End
   Begin VB.TextBox TxtEmpName 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   210
      TabIndex        =   3
      Top             =   4320
      Width           =   6660
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   1560
      Picture         =   "LOGIN.frx":1727FB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2790
      Width           =   855
   End
   Begin VB.CommandButton CmdContinue 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   540
      Picture         =   "LOGIN.frx":172B05
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2790
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10590
      Top             =   30
   End
   Begin MSMask.MaskEdBox TxtEmpNumber 
      Height          =   765
      Left            =   180
      TabIndex        =   0
      Top             =   1650
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1349
      _Version        =   393216
      BackColor       =   14737632
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label lblPMOUT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PM OUT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   6180
      TabIndex        =   33
      Top             =   3840
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblPMIN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PM IN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   5370
      TabIndex        =   32
      Top             =   3840
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lblAMOUT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AM OUT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4020
      TabIndex        =   31
      Top             =   3840
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblAMIN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AM IN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3240
      TabIndex        =   30
      Top             =   3840
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lblS 
      BackStyle       =   0  'Transparent
      Caption         =   "Key In Your Employee No. then press enter. Press again Key To Validate."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   240
      TabIndex        =   27
      Top             =   150
      Width           =   2535
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   0
      Left            =   13605
      Shape           =   3  'Circle
      Top             =   3045
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   1
      Left            =   13785
      Shape           =   3  'Circle
      Top             =   3045
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   2
      Left            =   14010
      Shape           =   3  'Circle
      Top             =   3060
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   3
      Left            =   14295
      Shape           =   3  'Circle
      Top             =   3075
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   4
      Left            =   14520
      Shape           =   3  'Circle
      Top             =   3060
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   5
      Left            =   14700
      Shape           =   3  'Circle
      Top             =   3045
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   6
      Left            =   14940
      Shape           =   3  'Circle
      Top             =   3030
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   7
      Left            =   15225
      Shape           =   3  'Circle
      Top             =   3030
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   8
      Left            =   15435
      Shape           =   3  'Circle
      Top             =   3015
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   9
      Left            =   15615
      Shape           =   3  'Circle
      Top             =   3045
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   10
      Left            =   15780
      Shape           =   3  'Circle
      Top             =   3030
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   11
      Left            =   15960
      Shape           =   3  'Circle
      Top             =   3015
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   12
      Left            =   16110
      Shape           =   3  'Circle
      Top             =   3015
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   13
      Left            =   16350
      Shape           =   3  'Circle
      Top             =   3030
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   14
      Left            =   16560
      Shape           =   3  'Circle
      Top             =   3030
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   15
      Left            =   16740
      Shape           =   3  'Circle
      Top             =   3030
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   16
      Left            =   16875
      Shape           =   3  'Circle
      Top             =   3030
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   17
      Left            =   17025
      Shape           =   3  'Circle
      Top             =   3030
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   18
      Left            =   17415
      Shape           =   3  'Circle
      Top             =   3030
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   19
      Left            =   17580
      Shape           =   3  'Circle
      Top             =   3030
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   20
      Left            =   17745
      Shape           =   3  'Circle
      Top             =   3030
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   21
      Left            =   17940
      Shape           =   3  'Circle
      Top             =   3045
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   22
      Left            =   18135
      Shape           =   3  'Circle
      Top             =   3030
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   23
      Left            =   12345
      Shape           =   3  'Circle
      Top             =   2835
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   24
      Left            =   12495
      Shape           =   3  'Circle
      Top             =   2820
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   25
      Left            =   12660
      Shape           =   3  'Circle
      Top             =   2835
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   26
      Left            =   12855
      Shape           =   3  'Circle
      Top             =   2805
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   27
      Left            =   13020
      Shape           =   3  'Circle
      Top             =   2805
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   28
      Left            =   13170
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   29
      Left            =   13350
      Shape           =   3  'Circle
      Top             =   2775
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   30
      Left            =   13530
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   31
      Left            =   13710
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   32
      Left            =   13905
      Shape           =   3  'Circle
      Top             =   2730
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   33
      Left            =   14100
      Shape           =   3  'Circle
      Top             =   2730
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   34
      Left            =   14265
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   35
      Left            =   14430
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   36
      Left            =   14610
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   37
      Left            =   14820
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   38
      Left            =   15000
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   39
      Left            =   15165
      Shape           =   3  'Circle
      Top             =   2775
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   40
      Left            =   15315
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   41
      Left            =   15465
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   42
      Left            =   15630
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   43
      Left            =   15795
      Shape           =   3  'Circle
      Top             =   2775
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   44
      Left            =   15945
      Shape           =   3  'Circle
      Top             =   2775
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   45
      Left            =   16095
      Shape           =   3  'Circle
      Top             =   2775
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   46
      Left            =   16245
      Shape           =   3  'Circle
      Top             =   2775
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   47
      Left            =   16395
      Shape           =   3  'Circle
      Top             =   2775
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   48
      Left            =   16545
      Shape           =   3  'Circle
      Top             =   2775
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   49
      Left            =   16695
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   50
      Left            =   16830
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   51
      Left            =   16965
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   52
      Left            =   17100
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   53
      Left            =   17235
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   54
      Left            =   17400
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   55
      Left            =   17550
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   150
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   56
      Left            =   17700
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   57
      Left            =   17850
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   58
      Left            =   18015
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   75
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   59
      Left            =   18150
      Shape           =   3  'Circle
      Top             =   2790
      Width           =   75
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   9
      Index           =   0
      X1              =   822
      X2              =   904
      Y1              =   229
      Y2              =   229
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      Index           =   1
      X1              =   930
      X2              =   1012
      Y1              =   229
      Y2              =   229
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   7
      Index           =   0
      X1              =   820
      X2              =   905
      Y1              =   246
      Y2              =   246
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   928
      X2              =   1013
      Y1              =   246
      Y2              =   246
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      Index           =   0
      X1              =   822
      X2              =   904
      Y1              =   261
      Y2              =   261
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   930
      X2              =   1012
      Y1              =   261
      Y2              =   261
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   9600
      Shape           =   3  'Circle
      Top             =   3600
      Width           =   165
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   26
      Left            =   16350
      MouseIcon       =   "LOGIN.frx":172E0F
      MousePointer    =   99  'Custom
      Top             =   3735
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   25
      Left            =   15840
      MouseIcon       =   "LOGIN.frx":172F61
      MousePointer    =   99  'Custom
      Top             =   3735
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   225
      Index           =   24
      Left            =   9570
      MousePointer    =   15  'Size All
      ToolTipText     =   "Move"
      Top             =   3570
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   23
      Left            =   16350
      MouseIcon       =   "LOGIN.frx":1730B3
      MousePointer    =   99  'Custom
      ToolTipText     =   "Exit"
      Top             =   3465
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   22
      Left            =   15840
      MouseIcon       =   "LOGIN.frx":173205
      MousePointer    =   99  'Custom
      ToolTipText     =   "Hide For 10 Seconds"
      Top             =   3465
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   150
      Index           =   21
      Left            =   16110
      MouseIcon       =   "LOGIN.frx":173357
      MousePointer    =   99  'Custom
      ToolTipText     =   "Help"
      Top             =   3285
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   585
      Index           =   20
      Left            =   11685
      Picture         =   "LOGIN.frx":1734A9
      Top             =   8100
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   585
      Index           =   19
      Left            =   7170
      Picture         =   "LOGIN.frx":1739CB
      Top             =   8100
      Width           =   180
   End
   Begin VB.Image Image2 
      Height          =   585
      Index           =   18
      Left            =   10515
      Picture         =   "LOGIN.frx":173F89
      ToolTipText     =   "Help"
      Top             =   8100
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   585
      Index           =   17
      Left            =   9855
      Picture         =   "LOGIN.frx":1744AB
      ToolTipText     =   "Swap Day And Month"
      Top             =   8100
      Width           =   150
   End
   Begin VB.Image Image2 
      Height          =   585
      Index           =   16
      Left            =   9090
      Picture         =   "LOGIN.frx":1749CD
      ToolTipText     =   "Move"
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   585
      Index           =   15
      Left            =   8475
      Picture         =   "LOGIN.frx":1751FB
      Top             =   8100
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   585
      Index           =   14
      Left            =   7860
      Picture         =   "LOGIN.frx":1755E5
      Top             =   8100
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   13
      Left            =   11430
      ToolTipText     =   "Year"
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   12
      Left            =   11175
      ToolTipText     =   "Year"
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   11
      Left            =   10920
      ToolTipText     =   "Year"
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   10
      Left            =   10665
      ToolTipText     =   "Year"
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   9
      Left            =   10260
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   8
      Left            =   10005
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   7
      Left            =   9600
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   6
      Left            =   9345
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   5
      Left            =   8835
      ToolTipText     =   "Second"
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   4
      Left            =   8580
      ToolTipText     =   "Second"
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   3
      Left            =   8220
      ToolTipText     =   "Minute"
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   2
      Left            =   7965
      ToolTipText     =   "Minute"
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   1
      Left            =   7605
      ToolTipText     =   "Hour"
      Top             =   8100
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   555
      Index           =   0
      Left            =   7350
      ToolTipText     =   "Hour"
      Top             =   8100
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp#"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   300
      TabIndex        =   19
      Top             =   5250
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   285
      TabIndex        =   18
      Top             =   5250
      Width           =   2265
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "IN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3885
      TabIndex        =   17
      Top             =   5250
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "OUT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5475
      TabIndex        =   16
      Top             =   5250
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   7110
      TabIndex        =   15
      Top             =   9030
      Width           =   4755
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   9570
      Shape           =   3  'Circle
      Top             =   3570
      Width           =   225
   End
   Begin VB.Image imgDispPic 
      Height          =   3525
      Left            =   3300
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Picture1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3585
      Left            =   3270
      Stretch         =   -1  'True
      Top             =   210
      Width           =   3675
   End
End
Attribute VB_Name = "frmLOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RESPONSE                                       As Integer
Private CRITERIA                                       As String
Dim CURRENTDATE                                        As Date
Dim DELAYCOUNTER                                       As Long
Dim HR                                                 As Integer
Dim MN                                                 As Integer
Dim SC                                                 As Integer
Dim PREVFOUND                                          As Boolean
Dim AMER_EURO                                          As Boolean
Dim AMPM                                               As String
Dim VAREMPNO                                           As String
Dim TIME_IN_AM                                         As String
Dim TIME_OUT_AM                                        As String
Dim TIME_IN_PM                                         As String
Dim TIME_OUT_PM                                        As String
Const PI = 3.141592654
Dim CURRENTTIME

Private Sub Calendar1_Click()
    Date = Calendar1.Value
End Sub

Sub CheckTimer()
    SC = Second(Time)
    If Minute(Time) < 55 Then
        If (Minute(Time) - MN) > 3 Then
            MN = Minute(Time)
            SC = Second(Time)
        End If
    Else
        Call Showtime
    End If
End Sub

Private Sub cmdCancel_Click()
    If TxtEmpNumber.Text <> "" Then
        RESPONSE = MsgBox("Do you really want to cancel this entry?", 4)
        If RESPONSE = 6 Then
            TxtEmpNumber.Text = ""
            TxtEmpName.Text = ""
            TxtTime.Text = ""
            TxtEmpNumber.SetFocus
            imgDispPic.Picture = LoadPicture("")
        Else
            CmdContinue.SetFocus
        End If
    Else
        TxtEmpNumber.SetFocus
    End If
End Sub

Private Sub CmdContinue_Click()
    
    Dim rsTmp                                          As New ADODB.Recordset
    Dim GREET                                          As String

    If TxtEmpNumber = "" Then
        MsgBox "Sorry, Employee number not yet selected", vbInformation
        TxtEmpNumber.SetFocus
        Exit Sub
    End If

   
    If Hour(TxtTime.Text) >= 0 And Hour(TxtTime.Text) < 12 Then
        GREET = "Good Morning"
    ElseIf Hour(TxtTime.Text) >= 12 And Hour(TxtTime.Text) < 18 Then
        GREET = "Good Afternoon"
    ElseIf Hour(TxtTime.Text) >= 18 And Hour(TxtTime.Text) < 24 Then
        GREET = "Good Evening"
    End If

    CRITERIA = "EMPNO = '" & VAREMPNO & "' AND DATETODAY = '" & Format(Now, "short date") & "'"

    If GetTimeShift(VAREMPNO) = 0 Then
        If TimeValue(TxtTime.Text) > TimeValue(#6:00:00 AM#) And TimeValue(TxtTime.Text) < TimeValue(#11:59:59 PM#) Then
            If TIME_IN_AM = "" Then
                GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET INAM = '" & Now & "' WHERE " & CRITERIA)
                Call CSMS_UPDATE_JOBSSTATUS("INAM", VAREMPNO)
                MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged In ", vbInformation
            ElseIf TIME_IN_AM <> "" And (DateDiff("n", TIME_IN_AM, Now) < 30) And TIME_OUT_PM = "" Then
                If MsgBox("" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you were already Logged In, Do you want to continue continue logging in again? ", vbYesNo, "") = vbYes Then
                    GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET INAM = '" & Now & "' WHERE " & CRITERIA)
                    Call CSMS_UPDATE_JOBSSTATUS("INAM", VAREMPNO)
                    MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged In ", vbInformation
                End If
            ElseIf TIME_IN_AM <> "" And TIME_OUT_PM = "" Then
                GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET OUTPM = '" & Now & "' WHERE " & CRITERIA)
                Call HRMS_UPDATE_ATTEND(VAREMPNO)
                Call CSMS_UPDATE_JOBSSTATUS("OUTPM", VAREMPNO)
                MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged Out ", vbInformation
            ElseIf TIME_IN_AM <> "" And TIME_OUT_PM <> "" And (DateDiff("n", TIME_OUT_PM, Now) < 30) Then
                If MsgBox("" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you were already Logged Out, Do you want to continue continue logging out again? ", vbYesNo, "") = vbYes Then
                    GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET OUTPM = '" & Now & "' WHERE " & CRITERIA)
                    Call HRMS_UPDATE_ATTEND(VAREMPNO)
                    Call CSMS_UPDATE_JOBSSTATUS("OUTPM", VAREMPNO)
                    MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged Out ", vbInformation
                End If
            Else
                MsgBox "Duplicate Entry Not Allowed", vbInformation
            End If
        Else
            Set rsTmp = New ADODB.Recordset
            Set rsTmp = GCONDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = '" & VAREMPNO & "' AND DATETODAY = '" & DateAdd("d", -1, Format(Now, "short date")) & "' AND OUTPM IS NULL")
            If Not rsTmp.EOF And Not rsTmp.BOF Then
                GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET OUTPM = '" & Now & "' WHERE EMPNO = '" & VAREMPNO & "' AND DATETODAY = '" & DateAdd("d", -1, Format(Now, "short date")) & "'")
                MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged Out ", vbInformation
            Else
                MsgBox "Duplicate Entry Not Allowed", vbInformation
            End If
        End If
    Else
        If TimeValue(TxtTime.Text) > TimeValue(#6:00:00 AM#) And TimeValue(TxtTime.Text) < TimeValue(#11:59:59 PM#) Then
            If TIME_IN_AM = "" Then
                GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET INAM = '" & Now & "' WHERE " & CRITERIA)
                Call CSMS_UPDATE_JOBSSTATUS("INAM", VAREMPNO)
                MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged In ", vbInformation
            ElseIf TIME_IN_AM <> "" And (DateDiff("n", TIME_IN_AM, Now) < 30) And TIME_OUT_PM = "" Then
                If MsgBox("" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you were already Logged In, Do you want to continue continue logging in again? ", vbYesNo, "") = vbYes Then
                    GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET INAM = '" & Now & "' WHERE " & CRITERIA)
                    Call CSMS_UPDATE_JOBSSTATUS("INAM", VAREMPNO)
                    MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged In ", vbInformation
                End If
            ElseIf TIME_IN_AM <> "" And TIME_OUT_AM = "" Then
                GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET OUTAM = '" & Now & "' WHERE " & CRITERIA)
                Call CSMS_UPDATE_JOBSSTATUS("OUTAM", VAREMPNO)
                MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged Out ", vbInformation
            ElseIf TIME_IN_AM <> "" And TIME_OUT_AM <> "" And (DateDiff("n", TIME_OUT_AM, Now) < 30) Then
                If MsgBox("" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you were already Logged Out, Do you want to continue continue logging out again? ", vbYesNo, "") = vbYes Then
                    GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET OUTAM = '" & Now & "' WHERE " & CRITERIA)
                    Call CSMS_UPDATE_JOBSSTATUS("OUTAM", VAREMPNO)
                    MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged Out ", vbInformation
                End If
            ElseIf TIME_IN_AM <> "" And TIME_OUT_AM <> "" And TIME_IN_PM = "" Then
                GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET INPM = '" & Now & "' WHERE " & CRITERIA)
                Call CSMS_UPDATE_JOBSSTATUS("INPM", VAREMPNO)
                MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged In ", vbInformation
            ElseIf TIME_IN_AM <> "" And TIME_OUT_AM <> "" And TIME_IN_PM <> "" And (DateDiff("n", TIME_IN_PM, Now) < 30) Then
                If MsgBox("" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you were already Logged In, Do you want to continue continue logging in again? ", vbYesNo, "") = vbYes Then
                    GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET INPM = '" & Now & "' WHERE " & CRITERIA)
                    Call CSMS_UPDATE_JOBSSTATUS("INPM", VAREMPNO)
                    MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged In ", vbInformation
                End If
            ElseIf TIME_IN_AM <> "" And TIME_OUT_AM <> "" And TIME_IN_PM <> "" And TIME_OUT_PM = "" Then
                GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET OUTPM = '" & Now & "' WHERE " & CRITERIA)
                Call HRMS_UPDATE_ATTEND(VAREMPNO)
                Call CSMS_UPDATE_JOBSSTATUS("OUTPM", VAREMPNO)
                MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged Out ", vbInformation
            ElseIf TIME_IN_AM <> "" And TIME_OUT_AM <> "" And TIME_IN_PM <> "" And TIME_OUT_PM <> "" And (DateDiff("n", TIME_OUT_PM, Now) < 30) Then
                If MsgBox("" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you were already Logged Out, Do you want to continue continue logging Out again? ", vbYesNo, "") = vbYes Then
                    GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET OUTPM = '" & Now & "' WHERE " & CRITERIA)
                    Call HRMS_UPDATE_ATTEND(VAREMPNO)
                    Call CSMS_UPDATE_JOBSSTATUS("OUTPM", VAREMPNO)
                    MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged Out ", vbInformation
                End If
            Else
                MsgBox "Duplicate Entry Not Allowed", vbInformation
            End If
        Else
            Set rsTmp = New ADODB.Recordset
            Set rsTmp = GCONDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = '" & VAREMPNO & "' AND DATETODAY = '" & DateAdd("d", -1, Format(Now, "short date")) & "' AND OUTPM IS NULL")
            If Not rsTmp.EOF And Not rsTmp.BOF Then
                GCONDMIS.Execute ("UPDATE HRMS_ATTEND SET OUTPM = '" & Now & "' WHERE EMPNO = '" & VAREMPNO & "' AND DATETODAY = '" & DateAdd("d", -1, Format(Now, "short date")) & "'")
                MsgBox "" & GREET & " : " & FindEmployeeName(VAREMPNO) & " you are now Logged Out ", vbInformation
            Else
                MsgBox "Duplicate Entry Not Allowed", vbInformation
            End If
        End If
    End If
    imgDispPic.Picture = LoadPicture("")
    TxtEmpNumber.Text = ""
    TxtEmpName.Text = ""
    TxtTime.Text = ""
    TxtEmpNumber.SetFocus
End Sub

 
Sub CSMS_UPDATE_JOBSSTATUS(INOUT As String, TechnicianID As String)

    Dim XTDATE                                         As Date
    XTDATE = Format(Date, "mm/dd/yy hh:mm:ss AM/PM")

    Dim TEMPRS                                         As ADODB.Recordset
    Set TEMPRS = GCONDMIS.Execute("select IS_TECHNICIAN, ISNULL(IS_TECHNICIAN,0), EMPNO from HRMS_EMPINFO WHERE EMPNO = '" & TechnicianID & "'")
    If Not TEMPRS.EOF And Not TEMPRS.BOF Then
        If Null2String(TEMPRS!is_technician) = False Then
            Exit Sub
        End If
    End If
    If TEMPRS.EOF Or TEMPRS.BOF Then
        Exit Sub
    End If

    If INOUT = "INAM" Then
        GCONDMIS.Execute " Update CSMS_JOBCLOCK SET Time_in_am='" & Time & "' WHERE TranDate= '" & XTDATE & "' AND  Technician= '" & TechnicianID & "'"
    ElseIf INOUT = "OUTAM" Then
        GCONDMIS.Execute " Update CSMS_JOBCLOCK SET Time_out_am ='" & Time & "' WHERE TranDate= '" & XTDATE & "' AND Technician='" & TechnicianID & "'"
    ElseIf INOUT = "INPM" Then
        GCONDMIS.Execute " Update CSMS_JOBCLOCK SET Time_in_PM='" & Time & "' WHERE TranDate='" & XTDATE & "' AND Technician='" & TechnicianID & "'"
    ElseIf INOUT = "OUTPM" Then
        GCONDMIS.Execute " Update CSMS_JOBCLOCK SET Time_out_PM='" & Time & "' WHERE TranDate='" & XTDATE & "' AND Technician='" & TechnicianID & "'"

        Dim PATS_NumHrs
        Dim PATS_STDHRS
        Set TEMPRS = GCONDMIS.Execute("SELECT ISNULL(TOTALHRSAM,0)TOTALHRSAM , ISNULL(TOTALHRSPM ,0) TOTALHRSPM , ISNULL(ActualHrsAM ,0)ActualHrsAM , ISNULL(ActualHrsPM, 0) ActualHrsPM FROM HRMS_ATTEND WHERE EMPNO='" & TechnicianID & "' AND DATETODAY= '" & XTDATE & "'")

        If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
            PATS_NumHrs = Round(TEMPRS!TotalHrsAm + TEMPRS!TotalHrsPm, 2)
            PATS_STDHRS = Round(TEMPRS!ActualHrsAm + TEMPRS!ActualHrsPm, 2)

            GCONDMIS.Execute _
                  " Update CSMS_JOBCLOCK SET NumHrs =" & PATS_NumHrs & _
                                                       ",  StdHrs = " & PATS_STDHRS & _
                                                     "  WHERE Technician='" & TechnicianID & "' AND TranDate= '" & XTDATE & "'"
        End If
        Set TEMPRS = Nothing
    End If
End Sub

Function FindEmployeeName(EMPCODE As String)
    Dim rsTmp                                          As New ADODB.Recordset
    Set rsTmp = GCONDMIS.Execute("SELECT FIRSTNAME,LASTNAME FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPCODE & "'")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        FindEmployeeName = N2String(rsTmp!LastName) & ", " & Null2String(rsTmp!FirstName)
    End If
    Set rsTmp = Nothing
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        If MsgBox("Close Personnel Attendance Tracking System?", vbQuestion + vbYesNo, "Closing PATS...") = vbYes Then
            Unload Me
            End
        End If
    End If
End Sub

Private Sub Form_Load()
    PREVFOUND = False
    If App.PrevInstance Then
        MsgBox "Previous Application already loaded," & Chr(13) & Chr(13) & _
               "Please double click the icon at the task bar.", vbInformation
    End If
    On Error Resume Next
    Dim rs                                             As ADODB.Recordset
    Set rs = GCONDMIS.Execute("SELECT getdate() as DateNow, host_name() as PCName")

    If rs!pcname <> "HARI-ADMIS" Then
        Date = rs!datenow
        Time = rs!datenow
        LOGTIME = Time
        LOGDATE = Date
    End If
    Set rs = Nothing
    CenterMe Screen, Me, 0
    DrawXPCtl Me
    Showtime
    Calendar1.Value = Now
    picOT.Visible = False
    Flag = -1
    TxtDate.Text = Format(Date, "dddddd")
    RefreshCurrentDate
End Sub

Sub GetEmployeeLogIn(vEMPNO As String)
    Dim rsTmp                                          As New ADODB.Recordset
    Set rsTmp = GCONDMIS.Execute("SELECT EMPNO, DATETODAY, INAM, OUTAM, INPM, OUTPM FROM HRMS_ATTEND WHERE EMPNO = '" & vEMPNO & "' AND DATETODAY = '" & Date & "'")
    If Not (rsTmp.EOF And rsTmp.BOF) Then
        TIME_IN_AM = Null2String(rsTmp!InAm)
        TIME_OUT_AM = Null2String(rsTmp!OutAm)
        TIME_IN_PM = Null2String(rsTmp!InPm)
        TIME_OUT_PM = Null2String(rsTmp!OutPM)
    Else
        TIME_IN_AM = ""
        TIME_OUT_AM = ""
        TIME_IN_PM = ""
        TIME_OUT_PM = ""
    End If
    Set rsTmp = Nothing
End Sub

Function GetTimeShift(vEMPNO As String) As Integer
    Dim rsTmp                                          As New ADODB.Recordset
    Dim RSSHIFT                                        As New ADODB.Recordset
    Set rsTmp = GCONDMIS.Execute("SELECT SHIFT FROM HRMS_EMPINFO WHERE EMPNO = '" & vEMPNO & "'")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Set RSSHIFT = GCONDMIS.Execute("SELECT PATS FROM HRMS_TIME_SHIFT_CODE WHERE SHIFTCODE = '" & rsTmp!Shift & "'")
        If Not (RSSHIFT.BOF And RSSHIFT.EOF) Then
            GetTimeShift = N2Str2Zero(RSSHIFT!PATS)
        End If
        Set RSSHIFT = Nothing
    End If
    Set rsTmp = Nothing
End Function

Sub HRMS_UPDATE_ATTEND(EmployeeID As String)
    Dim XTDATE                                         As Date
    XTDATE = Format(Date, "mm/dd/yy hh:mm:ss AM/PM")

    Dim totAM                                          As Double
    Dim totPM                                          As Double
    totAM = 0
    totPM = 0

    Dim rsTemp                                         As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = GCONDMIS.Execute("SELECT EMPNO, DATETODAY, INAM, OUTAM, INPM,OUTPM FROM HRMS_ATTEND WHERE EMPNO='" & EmployeeID & "' AND DATETODAY= '" & XTDATE & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        If GetTimeShift(TxtEmpNumber) = 0 Then
            totAM = Round((DateDiff("n", Format(rsTemp!InAm, "HH:MM:SS AM/PM"), #12:00:00 PM#)) \ 60 + Round(((DateDiff("n", Format(rsTemp!InAm, "HH:MM:SS AM/PM"), #12:00:00 PM#)) Mod 60) / 60, 2), 2)
            If totAM < 0 Then
                totAM = 0
            End If
            totPM = Round((DateDiff("n", #1:00:00 PM#, Format(rsTemp!OutPM, "HH:MM:SS AM/PM"))) \ 60 + Round(((DateDiff("n", #1:00:00 PM#, Format(rsTemp!OutPM, "HH:MM:SS AM/PM"))) Mod 60) / 60, 2), 2)
            If totPM < -1 Then
                totPM = totPM + 1
            ElseIf totPM < 0 Then
                totPM = Round(((DateDiff("n", #1:00:00 PM#, Format(rsTemp!OutPM, "HH:MM:SS AM/PM"))) Mod 60) / 60, 2)
            End If
        Else
            totAM = Round((DateDiff("n", rsTemp!InAm, rsTemp!OutAm)) \ 60 + Round(((DateDiff("n", rsTemp!InAm, rsTemp!OutAm)) Mod 60) / 60, 2), 2)
            totPM = Round((DateDiff("n", rsTemp!InPm, rsTemp!OutPM)) \ 60 + Round(((DateDiff("n", rsTemp!InPm, rsTemp!OutPM)) Mod 60) / 60, 2), 2)
        End If
    End If
    GCONDMIS.Execute "UPDATE HRMS_ATTEND SET TOTALHRSAM = " & totAM & ", " & _
                     "TOTALHRSPM = " & totPM & _
                   " WHERE EMPNO='" & EmployeeID & "' AND DATETODAY= '" & XTDATE & "'"
End Sub

Private Sub LOG()

End Sub

Sub RefreshCurrentDate()
    Dim rsEmpInfo2                                     As ADODB.Recordset
    Dim rsAttend2                                      As ADODB.Recordset
    Dim RSSHIFT                                        As ADODB.Recordset
    Dim vSHIFTINAM                                     As String
    Dim vSHIFTOUTAM                                    As String
    Dim vSHIFTINPM                                     As String
    Dim vSHIFTOUTPM                                    As String
    Dim vActualHrsAM                                   As String
    Dim vActualHrsPM                                   As Double
    vSHIFTINAM = ""
    vSHIFTOUTAM = ""
    vSHIFTINPM = ""
    vSHIFTOUTPM = ""

    Set rsEmpInfo2 = New ADODB.Recordset
    rsEmpInfo2.Open "SELECT EMPNO, SHIFT FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A' ORDER BY LASTNAME,FIRSTNAME,MIDDLENAME", GCONDMIS, adOpenKeyset
    If Not rsEmpInfo2.EOF And Not rsEMPINFO.BOF Then
        rsEMPINFO.MoveFirst
        Do While Not rsEmpInfo2.EOF
            Set RSSHIFT = New ADODB.Recordset
            Set RSSHIFT = GCONDMIS.Execute("SELECT * FROM HRMS_TIME_SHIFT_CODE WHERE SHIFTCODE = " & N2Str2Null(rsEmpInfo2!Shift))
            If Not RSSHIFT.EOF And Not RSSHIFT.BOF Then
                If Null2String(RSSHIFT!from1) = "" Then
                    vSHIFTINAM = Format("08:00:00 AM", "HH:MM")
                End If
                If Not Null2String(RSSHIFT!from1) = "" Then
                    vSHIFTINAM = Format(Null2String(RSSHIFT!from1), "HH:MM")
                End If
                If Null2String(RSSHIFT!lunchout) = "" Then
                    vSHIFTOUTAM = Format("12:00:00 PM", "HH:MM")
                End If
                If Not Null2String(RSSHIFT!lunchout) = "" Then
                    vSHIFTOUTAM = Format(Null2String(RSSHIFT!lunchout), "HH:MM")
                End If
                If Null2String(RSSHIFT!lunchin) = "" Then
                    vSHIFTINPM = Format("01:00:00 PM", "HH:MM")
                End If
                If Not Null2String(RSSHIFT!lunchin) = "" Then
                    vSHIFTINPM = Format(Null2String(RSSHIFT!lunchin), "HH:MM")
                End If
                If Null2String(RSSHIFT!to1) = "" Then
                    vSHIFTOUTPM = Format("05:00:00 PM", "HH:MM")
                End If
                If Not Null2String(RSSHIFT!to1) = "" Then
                    vSHIFTOUTPM = Format(Null2String(RSSHIFT!to1), "HH:MM")
                End If
                vActualHrsAM = DateDiff("h", vSHIFTINAM, vSHIFTOUTAM)
                vActualHrsPM = DateDiff("h", vSHIFTINPM, vSHIFTOUTPM)
            Else
                vSHIFTINAM = Format("08:00:00 AM", "HH:MM")
                vSHIFTOUTAM = Format("12:00:00 PM", "HH:MM")
                vSHIFTINPM = Format("01:00:00 PM", "HH:MM")
                vSHIFTOUTPM = Format("05:00:00 PM", "HH:MM")

                vActualHrsAM = DateDiff("h", vSHIFTINAM, vSHIFTOUTAM)
                vActualHrsPM = DateDiff("h", vSHIFTINPM, vSHIFTOUTPM)
            End If
            Set rsAttend2 = GCONDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE DATETODAY = '" & Format(Date, "Short Date") & "' AND EMPNO = " & N2Str2Null(rsEmpInfo2!empno))
            If (rsAttend2.EOF And rsAttend2.BOF) Then
                GCONDMIS.Execute "INSERT INTO HRMS_ATTEND (EMPNO, SHIFT, SHIFTINAM, SHIFTOUTAM, SHIFTINPM, SHIFTOUTPM, DATETODAY, ACTUALHRSAM, ACTUALHRSPM) " & _
                               " Values(" & N2Str2Null(rsEmpInfo2!empno) & _
                                 "," & N2Str2Null(rsEmpInfo2!Shift) & _
                                 "," & N2Str2Null(vSHIFTINAM) & _
                                 "," & N2Str2Null(vSHIFTOUTAM) & _
                                 "," & N2Str2Null(vSHIFTINPM) & _
                                 "," & N2Str2Null(vSHIFTOUTPM) & _
                                 ",'" & Date & _
                                 "'," & vActualHrsAM & _
                                 "," & vActualHrsPM & ")"
            End If
            rsEmpInfo2.MoveNext
        Loop
    End If
End Sub

Sub ShowClock()
    Dim varBaseX, varBaseY, I                          As Integer
    varBaseX = 645
    varBaseY = 245
    For I = 0 To 59
        If I Mod 5 = 0 Then
            Shape4(I).Left = varBaseX + Cos(I * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
            Shape4(I).Top = varBaseY + Sin(I * 2 * PI / 60 - (0.5 * PI)) * 100 - 5
        Else
            Shape4(I).Left = varBaseX + Cos(I * 2 * PI / 60 - (0.5 * PI)) * 100 - 2.5
            Shape4(I).Top = varBaseY + Sin(I * 2 * PI / 60 - (0.5 * PI)) * 100 - 2.5
        End If
        Shape4(I).BorderColor = vbBlack
    Next I
End Sub

Function Showtime()
    HR = Hour(Time)
    If HR = 0 Then
        HR = 12
    End If
    If HR >= 12 Then
        AMPM = "PM"
    Else
        AMPM = "AM"
    End If
    If HR > 12 Then
        HR = HR - 12
    End If
    MN = Minute(Time)
    SC = Second(Time)
    SC = SC + 1
    If SC = 60 Then
        SC = 0
        MN = MN + 1
        If MN = 60 Then
            MN = 0
            HR = HR + 1
            If HR = 12 Then
                If AMPM = "PM" Then
                    AMPM = "AM"
                Else
                    AMPM = "PM"
                End If
            End If
            If HR = 13 Then
                HR = 1
            End If
        End If
    End If

    Showtime = Format(HR, "00") + ":" + Format(MN, "00") + ":" + Format(SC, "00") + " " + AMPM
    CURRENTTIME = Format(HR, "00") + ":" + Format(MN, "00") + ":" + Format(SC, "00") + " " + AMPM
    If Date <> CURRENTDATE Then
        Calendar1.Value = Date
        RefreshCurrentDate
        CURRENTDATE = Date
    End If
    If TxtDate.Text <> Date Then
        TxtDate.Text = Format(Date, "dddddd")
    End If
End Function

Private Sub Timer1_Timer()
    Call CheckTimer
    Label1.Caption = Showtime()
End Sub

Private Sub Timer2_Timer()
    Image2(0) = Image1(Mid$(Time$, 1, 1)): Image2(1) = Image1(Mid$(Time$, 2, 1)): Image2(2) = Image1(Mid$(Time$, 4, 1)): Image2(3) = Image1(Mid$(Time$, 5, 1)): Image2(4) = Image1(Mid$(Time$, 7, 1)): Image2(5) = Image1(Mid$(Time$, 8, 1))
    If AMER_EURO Then
        Image2(6) = Image1(Mid$(Date$, 1, 1)): Image2(6).ToolTipText = "Month"
        Image2(7) = Image1(Mid$(Date$, 2, 1)): Image2(7).ToolTipText = "Month"
        Image2(8) = Image1(Mid$(Date$, 4, 1)): Image2(8).ToolTipText = "Day"
        Image2(9) = Image1(Mid$(Date$, 5, 1)): Image2(9).ToolTipText = "Day"
    Else
        Image2(8) = Image1(Mid$(Date$, 1, 1)): Image2(8).ToolTipText = "Month"
        Image2(9) = Image1(Mid$(Date$, 2, 1)): Image2(9).ToolTipText = "Month"
        Image2(6) = Image1(Mid$(Date$, 4, 1)): Image2(6).ToolTipText = "Day"
        Image2(7) = Image1(Mid$(Date$, 5, 1)): Image2(7).ToolTipText = "Day"
    End If
    Image2(10) = Image1(Mid$(Date$, 7, 1)): Image2(11) = Image1(Mid$(Date$, 8, 1))
    Image2(12) = Image1(Mid$(Date$, 9, 1)): Image2(13) = Image1(Mid$(Date$, 10, 1))
End Sub

Private Sub Timer4_Timer()
    Dim I                                              As Long
    Dim varBaseX                                       As Integer
    Dim varBaseY                                       As Integer
    varBaseX = 645
    varBaseY = 245
    Dim Tim                                            As Long
    Tim = Int(Timer)
    For I = 0 To 1
        'Hour
        Line1(I).X1 = varBaseX: Line1(I).Y1 = varBaseY
        Line1(I).X2 = varBaseX + Cos((Tim Mod 43200) * 2 * PI / 43200 - (0.5 * PI)) * 60
        Line1(I).Y2 = varBaseY + Sin((Tim Mod 43200) * 2 * PI / 43200 - (0.5 * PI)) * 60
        'Minute
        Line2(I).X1 = varBaseX: Line2(I).Y1 = varBaseY
        Line2(I).X2 = varBaseX + Cos((Tim \ 60 Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 90
        Line2(I).Y2 = varBaseY + Sin((Tim \ 60 Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 90
        'Second
        Line3(I).X1 = varBaseX - Cos((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 15
        Line3(I).Y1 = varBaseY - Sin((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 15
        Line3(I).X2 = varBaseX + Cos((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 90
        Line3(I).Y2 = varBaseY + Sin((Tim Mod 60) * 2 * PI / 60 - (0.5 * PI)) * 90
    Next I
End Sub

Private Sub Timer5_Timer()
    DELAYCOUNTER = DELAYCOUNTER + 1
    If DELAYCOUNTER >= 100 Then
        ShowClock
        Timer5.Interval = 0
    End If
End Sub

Private Sub Timer6_Timer()
    If lblS.ForeColor = vbRed Then
        lblS.ForeColor = vbYellow
    Else
        lblS.ForeColor = vbRed
    End If
End Sub

Private Sub TxtEmpName_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub TxtEmpNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtEmpNumber_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtEmpNumber_Validate(Cancel As Boolean)
    If TxtEmpName = "" Then
    Cancel = True
    End If
    Set rsEMPINFO = New ADODB.Recordset
    'FOR HAI
    'rsEMPINFO.Open "SELECT * FROM HRMS_EMPINFO WHERE RIGHT(EMPNO,4) = '" & TxtEmpNumber.Text & "'", gconDMIS
    rsEMPINFO.Open "SELECT EMPNO, PICFILNAME, LASTNAME, FIRSTNAME, MIDDLENAME, ACTIVEINACTIVE FROM HRMS_EMPINFO WHERE EMPNO = '" & TxtEmpNumber.Text & "'", GCONDMIS
    If rsEMPINFO.EOF Then
        imgDispPic.Picture = LoadPicture("")
        TxtEmpName.Text = ""
        RESPONSE = MsgBox("Employee Number NOT FOUND", vbInformation)
        TxtEmpNumber.Text = ""
        TxtEmpNumber.SetFocus
        Exit Sub
    Else
        VAREMPNO = Null2String(rsEMPINFO!empno)
        If Null2String(rsEMPINFO!picfilname) <> "" Then
            On Error Resume Next
            LoadPic imgDispPic, HRMS_PICTURES_PATH & Null2String(rsEMPINFO!picfilname)
        Else
            LoadPic imgDispPic, ""
        End If
        TxtEmpName.Text = Null2String(rsEMPINFO!LastName) & "," & Null2String(rsEMPINFO!FirstName) & " " & Left(Null2String(rsEMPINFO!middlename), 1) & "."
        If rsEMPINFO!ACTIVEINACTIVE = "I" Then
            RESPONSE = MsgBox("Employee NOT ACTIVE", 0)
            TxtEmpNumber.Text = ""
            TxtEmpName.Text = ""
            imgDispPic.Picture = LoadPicture("")
            TxtEmpNumber.SetFocus
            Exit Sub
        End If
        Dim TempAttend                                 As ADODB.Recordset
        Set TempAttend = New ADODB.Recordset
        TempAttend.Open "SELECT EMPNO, DATETODAY FROM HRMS_ATTEND WHERE EMPNO = '" & VAREMPNO & "' AND DATETODAY = '" & Format(Now, "short date") & "'", GCONDMIS
        If Not TempAttend.EOF And Not TempAttend.BOF Then
            GetEmployeeLogIn VAREMPNO
        End If
        TxtTime.Text = Time
        TempAttend.Close
         
    End If
End Sub
