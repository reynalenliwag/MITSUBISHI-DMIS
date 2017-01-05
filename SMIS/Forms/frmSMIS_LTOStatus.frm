VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO50BF~1.OCX"
Begin VB.Form frmSMIS_LTOStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LTO Status"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSMIS_LTOStatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame lblMake 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2145
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7695
      Begin Crystal.CrystalReport LTORPT 
         Left            =   2190
         Top             =   390
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label lblAdd 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "zzzzz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   555
         Left            =   90
         TabIndex        =   2
         Top             =   600
         Width           =   7485
      End
      Begin VB.Label lblCustName 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxxxxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   7485
      End
      Begin VB.Label lblSalesAE 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxxxxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   4
         Top             =   1170
         Width           =   2745
      End
      Begin VB.Label lblVDR 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5340
         TabIndex        =   11
         Top             =   1485
         Width           =   2205
      End
      Begin VB.Label lblAgingDP 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   13
         Top             =   1830
         Width           =   2745
      End
      Begin VB.Label lblAgingDI 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   9
         Top             =   1500
         Width           =   2745
      End
      Begin VB.Label lblVI 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5340
         TabIndex        =   6
         Top             =   1185
         Width           =   2205
      End
      Begin VB.Label lblSOno 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "xxxxxxxxxxxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5340
         TabIndex        =   14
         Top             =   1815
         Width           =   2205
      End
      Begin VB.Label lblDaysInventory 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aging (PullOut)"
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
         Height          =   225
         Left            =   105
         TabIndex        =   12
         Top             =   1830
         Width           =   1260
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aging (Sold) "
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
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VDR NO."
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
         Height          =   225
         Index           =   19
         Left            =   4545
         TabIndex        =   10
         Top             =   1500
         Width           =   705
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VI NO."
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
         Height          =   225
         Index           =   18
         Left            =   4740
         TabIndex        =   5
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale AE"
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
         Height          =   225
         Index           =   16
         Left            =   105
         TabIndex        =   3
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VSO NO."
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
         Height          =   225
         Index           =   7
         Left            =   4530
         TabIndex        =   7
         Top             =   1830
         Width           =   720
      End
   End
   Begin VB.Frame frmVehInfo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1065
      Left            =   60
      TabIndex        =   19
      Top             =   2100
      Width           =   7695
      Begin VB.Label lblEngno 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   4680
         TabIndex        =   26
         Top             =   570
         Width           =   2265
      End
      Begin VB.Label lblVIN 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   4680
         TabIndex        =   23
         Top             =   210
         Width           =   2265
      End
      Begin VB.Label lblIgnKeyNo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   930
         TabIndex        =   27
         Top             =   600
         Width           =   2715
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   930
         TabIndex        =   21
         Top             =   210
         Width           =   2715
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color "
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
         Index           =   15
         Left            =   375
         TabIndex        =   20
         Top             =   270
         Width           =   495
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CS NO."
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
         Index           =   14
         Left            =   210
         TabIndex        =   25
         Top             =   630
         Width           =   615
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Engine no. "
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
         Index           =   5
         Left            =   3750
         TabIndex        =   24
         Top             =   630
         Width           =   930
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vin no. "
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
         Index           =   4
         Left            =   3960
         TabIndex        =   22
         Top             =   270
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CSR Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1845
      Left            =   120
      TabIndex        =   28
      Top             =   3150
      Width           =   7665
      Begin VB.TextBox txtcer 
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
         Left            =   1050
         TabIndex        =   54
         Top             =   1290
         Width           =   1965
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frmSMIS_LTOStatus.frx":08CA
         Left            =   3180
         List            =   "frmSMIS_LTOStatus.frx":08E3
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   570
         Width           =   4335
      End
      Begin VB.TextBox txtOthers 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3180
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   930
         Visible         =   0   'False
         Width           =   4305
      End
      Begin VB.TextBox txtPlateno 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1050
         MaxLength       =   6
         TabIndex        =   33
         Top             =   960
         Width           =   1965
      End
      Begin VB.TextBox txtCSRno 
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
         Left            =   1050
         TabIndex        =   32
         Top             =   660
         Width           =   1965
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   1020
         TabIndex        =   30
         Top             =   270
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830593
         CurrentDate     =   39336
      End
      Begin VB.Label Label1 
         Caption         =   "Certificate"
         Height          =   225
         Left            =   60
         TabIndex        =   53
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warranty "
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   52
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LTO Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3180
         TabIndex        =   41
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CSR NO."
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   31
         Top             =   660
         Width           =   810
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate no. "
         Height          =   210
         Index           =   6
         Left            =   90
         TabIndex        =   34
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CSR Date "
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   29
         ToolTipText     =   "Certificate Stock Reported Date From LTO Office"
         Top             =   330
         Width           =   900
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   -540
      ScaleHeight     =   915
      ScaleWidth      =   8910
      TabIndex        =   45
      Top             =   5820
      Width           =   8910
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   7590
         MouseIcon       =   "frmSMIS_LTOStatus.frx":0994
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_LTOStatus.frx":0AE6
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   6900
         MouseIcon       =   "frmSMIS_LTOStatus.frx":0E4C
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_LTOStatus.frx":0F9E
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   6210
         MouseIcon       =   "frmSMIS_LTOStatus.frx":1304
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_LTOStatus.frx":1456
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Fin&d"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   5520
         MouseIcon       =   "frmSMIS_LTOStatus.frx":17B2
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_LTOStatus.frx":1904
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4830
         MouseIcon       =   "frmSMIS_LTOStatus.frx":1BFE
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_LTOStatus.frx":1D50
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4140
         MouseIcon       =   "frmSMIS_LTOStatus.frx":20A8
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_LTOStatus.frx":21FA
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.Frame fraMRR 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6645
      Left            =   60
      TabIndex        =   15
      Top             =   90
      Visible         =   0   'False
      Width           =   7695
      Begin XtremeReportControl.ReportControl ReportControl1 
         Height          =   5775
         Left            =   60
         TabIndex        =   36
         Top             =   750
         Width           =   7575
         _Version        =   655364
         _ExtentX        =   13361
         _ExtentY        =   10186
         _StockProps     =   64
         BorderStyle     =   2
      End
      Begin VB.OptionButton optCust 
         BackColor       =   &H00808080&
         Caption         =   "Client"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   210
         Width           =   825
      End
      Begin VB.CommandButton cmdCancelSearch 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6990
         TabIndex        =   37
         ToolTipText     =   "Add Reminder"
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optCSNO 
         BackColor       =   &H00800000&
         Caption         =   "CS#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   210
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton optVI 
         BackColor       =   &H00008000&
         Caption         =   "VI#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   210
         Width           =   825
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2670
         MaxLength       =   10
         TabIndex        =   18
         Top             =   210
         Width           =   4275
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6300
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   42
      Top             =   5820
      Width           =   1800
      Begin VB.CommandButton cmdCancel 
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
         Height          =   795
         Left            =   750
         MouseIcon       =   "frmSMIS_LTOStatus.frx":2559
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_LTOStatus.frx":26AB
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
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
         Height          =   795
         Left            =   60
         MouseIcon       =   "frmSMIS_LTOStatus.frx":29E9
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_LTOStatus.frx":2B3B
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.Label LabID 
      Height          =   345
      Left            =   10290
      TabIndex        =   35
      Top             =   6240
      Width           =   495
   End
End
Attribute VB_Name = "frmSMIS_LTOStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CSNO                                                              As String
Dim rsLTO                                                             As New ADODB.Recordset

Dim LTO_ADD_EDIT                                                      As String

Function gettheDatereleased() As Date
    Dim RS                                                            As New ADODB.Recordset
    Dim tmp                                                           As Date
    Dim SQL                                                           As String

    SQL = "SELECT datereleased from SMIS_mrrinv_table where ignkey='" & lblIgnKeyNo & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    Do While Not RS.EOF
        On Error Resume Next
        gettheDatereleased = Null2String(RS!DateReleased)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Function

Function CheckIfPlateExist() As Boolean
    Dim RS                                                            As New ADODB.Recordset
    Dim SQL                                                           As String

    SQL = "SELECT plate_no from CSMS_cusveh where plate_no='" & txtPlateno & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        CheckIfPlateExist = True
    End If
    Set RS = Nothing
End Function

Function loadVehID() As Integer
    Dim RS                                                            As New ADODB.Recordset
    Dim SQL                                                           As String

    SQL = "SELECT ID,plate_no from CSMS_cusveh where Vcond_no='" & lblIgnKeyNo & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        loadVehID = Null2String(RS!ID)
        If Null2String(RS!PLATE_NO) = "" Then

        End If
    End If
    Set RS = Nothing
End Function

Sub UpdateStataus(MBNO)
    CSNO = MBNO
End Sub

Sub CleanAllLabel()
    lblSOno.Caption = ""
    lblAgingDI.Caption = ""
    lblAgingDP.Caption = ""
    lblVI.Caption = ""
    lblVDR.Caption = ""
    lblSalesAE.Caption = ""
    lblCustName.Caption = ""
    lblAdd.Caption = ""
    '    lblTelno.Caption = ""
    '    lblOffno.Caption = ""
    lblMake.Caption = ""
    lblColor.Caption = ""
    lblIgnKeyNo.Caption = ""
    lblVIN.Caption = ""
    lblEngno.Caption = ""

    txtPlateno.Text = ""
    '    txtCSRDate.Text = ""
    txtCSRno.Text = ""
End Sub

Sub initGrid()
    ReportControlPaintManager ReportControl1
    ReportControlAddColumnHeader ReportControl1, "CSNO, MODEL,PULLOUTDATE,VI_NO,CUSTOMERNAME, LTO "
    ResizeColumnHeader ReportControl1, "30,60,30,100,20,20,0"
    ReportControl1.GroupsOrder.Add ReportControl1.Columns(5)
    ReportControl1.Columns(5).Visible = False

End Sub

Sub StoreMemVars()
    If Not (rsLTO.BOF Or rsLTO.EOF) Then
        lblSOno.Caption = Null2String(rsLTO!SO_NO)
        labid = rsLTO!ID
        If IsDate(rsLTO!InvoicedDate) = True Then
            lblAgingDI.Caption = DateDiff("d", rsLTO!InvoicedDate, Date) & " Days " & rsLTO!InvoicedDate
            lblDaysInventory = "Inventory Days"
            If IsDate(rsLTO!PullOutDate) = True Then
                lblAgingDP.Caption = DateDiff("d", rsLTO!PullOutDate, rsLTO!InvoicedDate) & " Days " & rsLTO!PullOutDate
            Else
                lblAgingDP.Caption = ""
            End If
        Else
            lblAgingDI.Caption = ""
            lblDaysInventory = "Vehicle Aging"
            If IsDate(rsLTO!PullOutDate) = True Then
                lblAgingDP.Caption = DateDiff("d", rsLTO!PullOutDate, Date) & " Days " & rsLTO!PullOutDate
            Else
                lblAgingDP.Caption = ""
            End If
        End If

        lblVI.Caption = Null2String(rsLTO!VI_NO)

        If Null2String(rsLTO!VI_NO) = "" Then
            txtPlateno.Enabled = False
        Else
            txtPlateno.Enabled = True
        End If

        lblVDR.Caption = Null2String(rsLTO!VDR_NO)
        lblSalesAE.Caption = Null2String(rsLTO!salesae)
        lblCustName.Caption = Null2String(rsLTO!CustName)
        lblAdd.Caption = Null2String(rsLTO!HomeAddress)
        DTPicker1.Value = Null2Date(rsLTO!CSRDATE)
        lblMake.Caption = Null2String(rsLTO!DESCRIPT)
        lblColor.Caption = Null2String(rsLTO!Color)
        lblIgnKeyNo.Caption = Null2String(rsLTO!ignkey)
        lblVIN.Caption = Null2String(rsLTO!Vino)
        lblEngno.Caption = Null2String(rsLTO!EngineNo)
        txtPlateno.Text = Null2String(rsLTO!PLATE_NO)
        txtCSRno.Text = Null2String(rsLTO!CSR)
        txtcer.Text = Null2String(rsLTO!certific8)

        If Null2String(rsLTO!LTOStatus) = "" Then
            Combo1.ListIndex = -1
        Else
            Combo1.ListIndex = SelectCombo(Combo1, Null2String(rsLTO!LTOStatus))
            If Combo1.ListIndex = -1 Then: Combo1.ListIndex = 0
        End If

        loadVehID

    End If
End Sub

Sub rsRefresh()
    Dim SQL
    SQL = " SELECT "
    SQL = SQL & " A.ID, "
    SQL = SQL & " ISNULL(A.YEER ,'') + ISNULL(' ' + A.MAKE ,'') + ISNULL(' ' + A.DESCRIPT ,'') AS DESCRIPT ,"
    SQL = SQL & " A.COLOR, A.SOURCE, A.IGNKEY, A.ENGINENO, A.VINO, B.VDR_NO,CERTIFIC8, B.VI_NO, B.SO_NO, B.PLATE_NO,  "
    SQL = SQL & " A.DATERECEIVED, A.DATERELEASED, B.INVOICEDDATE, A.PULLOUTDATE,"
    SQL = SQL & " A.LTOSTATUS, A.CSR,A.CSRDATE, "
    SQL = SQL & " B.CUSTNAME , B.HOMEADDRESS , B.SALESAE"
    SQL = SQL & " FROM"
    SQL = SQL & " SMIS_MRRINV A LEFT OUTER JOIN"
    SQL = SQL & " SMIS_SALESORDER B ON A.IGNKEY = B.IGNKEY_NO"
    SQL = SQL & " WHERE A.STATUS ='P'"
    SQL = SQL & " AND A.VI_NO IS NOT NULL "
    SQL = SQL & " ORDER BY A.VI_NO"

    Call rsLTO.Open(SQL, gconDMIS, adOpenKeyset)
End Sub

Sub FillSearchGrid()
    'CSNO, MODEL,PULLOUTDATE,VI_NO,CUSTOMERNAME, LTO
    Dim SEARCHSTRING, SQL                                             As String
    If optVI.Value = True Then
        SEARCHSTRING = " WHERE SMIS_SALESORDER.VI_NO like '" & Format(txtSearch, "000000") & "%'"

    ElseIf optCust.Value = True Then
        SEARCHSTRING = " WHERE CUSTNAME like '%" & txtSearch & "%'"
    Else
        SEARCHSTRING = " WHERE IGNKEY like '" & txtSearch & "%'"
    End If
    'VI_NO,CUSTOMERNAME,CSNO, MODEL,PULLOUTDATE, LTO, STATUS

    SQL = "SELECT IGNKEY , YEER + ' ' + MAKE + ' '  + DESCRIPT  , PULLOUTDATE ,  SMIS_SALESORDER.VI_NO , CUSTNAME , LTOSTATUS FROM SMIS_MRRINV " & vbCrLf
    SQL = SQL & " LEFT OUTER JOIN SMIS_SALESORDER " & vbCrLf
    SQL = SQL & "  ON  IGNKEY=IGNKEY_NO " & SEARCHSTRING & " and  SMIS_SALESORDER.STATUS<>'C' ORDER BY LTOSTATUS, PULLoUTDATE DESC "

    flex_FillReportView gconDMIS.Execute(SQL), ReportControl1

End Sub

Private Sub cmdCancel_Click()
    picAdds.Visible = True
    picSaves.Visible = False
    Frame1.Enabled = False
End Sub

Private Sub cmdCancelSearch_Click()
    fraMRR.ZOrder 0
    fraMRR.Visible = False
End Sub

Private Sub cmdEdit_Click()
    LTO_ADD_EDIT = "EDIT"
    picAdds.Visible = False
    picSaves.Visible = True
    Frame1.Enabled = True

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    fraMRR.Visible = True
    fraMRR.ZOrder 0
    FillSearchGrid
End Sub

Private Sub cmdNext_Click()
    rsLTO.MoveNext
    If rsLTO.EOF Then
        rsLTO.MoveLast
        Call ShowLastRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdPrint_Click()

    Dim FILTER                                                        As String
    On Error GoTo ErrorCode:


    Screen.MousePointer = 11

    LTORPT.WindowTitle = "LTO STATUS REPORT"
    LTORPT.Reset
    LTORPT.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    LTORPT.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    LogAudit "V", "LTO STATUS ", "FOR " & lblVI
    NEW_LogAudit "V", "LTO STATUS", "", Null2String(labid), "", "CSR:" & N2Str2Null(lblIgnKeyNo), "", ""
    PrintSQLReport LTORPT, SMIS_REPORT_PATH & "LTOrep.rpt", "{SMIS_MRRINV_TABLE.ignkey}='" & lblIgnKeyNo & "'", DMIS_REPORT_Connection, 1

    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    Dim LTO, CSRNO                                                    As String
    If UCase(Combo1) = "OTHERS" Then
        LTO = txtOthers
    Else
        LTO = Combo1.Text
    End If

    If IsNull(DTPicker1.Value) = True Then
        CSRNO = N2Str2Null("")
    Else
        CSRNO = N2Str2Null(txtCSRno)

    End If
    If IsNull(DTPicker1.Value) = True Then
        SQL_STATEMENT = ("UPDATE SMIS_MRRINV SET LTOSTATUS=" & N2Str2Null(LTO) & " , CSR=" & N2Str2Null(CSRNO) & " , CSRDATE=null  WHERE ID=" & labid)
        NEW_LogAudit "E", "LTO STATUS", SQL_STATEMENT, Null2String(labid), "", "CSR:" & N2Str2Null(lblIgnKeyNo), "", ""
        gconDMIS.Execute SQL_STATEMENT
    Else
        SQL_STATEMENT = ("UPDATE SMIS_MRRINV SET LTOSTATUS=" & N2Str2Null(LTO) & " , CSR=" & N2Str2Null(CSRNO) & " , CSRDATE='" & DTPicker1.Value & "'  WHERE ID=" & labid)
        NEW_LogAudit "E", "LTO STATUS", SQL_STATEMENT, Null2String(labid), "", "CSR:" & N2Str2Null(lblIgnKeyNo), "", ""
        gconDMIS.Execute SQL_STATEMENT
        
    End If

    If lblVI <> Null Then
        'If CheckIfPlateExist = True Then
        '      MsgBox "yaon na"
        'End If
        gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET PLATE_NO=" & N2Str2Null(txtPlateno) & ",CERTIFIC8=" & N2Str2Null(txtcer) & " WHERE VI_NO='" & lblVI & "'")
        gconDMIS.Execute (" Update CSMS_CUSVEH SET  PLATE_NO='" & txtPlateno & "',d_sold= '" & gettheDatereleased & "',war_cert=" & N2Str2Null(txtcer) & " WHERE VCOND_NO=" & N2Str2Null(lblIgnKeyNo))
        LogAudit "E", "UPDATE CUSVEH/LTO-MONITORING", "PLATE NO " & txtPlateno
    Else
        LogAudit "E", "UPDATE LTO-MONITORING", "PLATE NO " & txtPlateno
        
        ShowSuccessFullyUpdated
        
    End If



    rsLTO.Requery
    rsLTO.Find ("ID=" & labid)
    StoreMemVars
    cmdCancel.Value = True
End Sub

Private Sub cmdPrevious_Click()
    rsLTO.MovePrevious
    If rsLTO.BOF Then
        rsLTO.MoveFirst
        Call ShowLastRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub Combo1_Click()
    If UCase(Combo1) = "OTHERS" Then
        txtOthers.Visible = True
    Else
        txtOthers.Visible = False
    End If
End Sub

Private Sub DTPicker1_Click()
    txtCSRno.Enabled = Not (IsNull(DTPicker1.Value))
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (LTO STATUS)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labid), "LTO STATUS")
            'End If
    End Select

End Sub

Private Sub Form_Load()
    Call rsRefresh
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    initGrid
    If CSNO <> "" Then
        rsLTO.Find ("IGNKEY='" & CSNO & "'")

    End If
    Frame1.Enabled = False
    Call StoreMemVars
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsLTO = Nothing
End Sub

Private Sub optCSNO_Click()
    txtSearch.SetFocus
End Sub

Private Sub optCust_Click()
    txtSearch.SetFocus
End Sub

Private Sub optVI_Click()
    txtSearch.SetFocus
End Sub

Private Sub ReportControl1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.GroupRow = True Then: Exit Sub
    rsLTO.MoveFirst
    rsLTO.Find ("IGNKEY='" & Row.Record(0).Value & "'")
    StoreMemVars
    cmdCANCELSEARCH.Value = True
    cmdEdit.Value = True
End Sub

Private Sub txtPlateno_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSearch_Change()
    FillSearchGrid
End Sub

