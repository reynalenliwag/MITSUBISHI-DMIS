VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCSMSEditCustomerVehicle_neo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Vehicle Maintenance"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSEditCustomerVehicle_neo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   14910
   Begin XtremeReportControl.ReportControl rptVEH 
      Height          =   6375
      Left            =   60
      TabIndex        =   80
      Top             =   780
      Width           =   14835
      _Version        =   655364
      _ExtentX        =   26167
      _ExtentY        =   11245
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnReorder=   0   'False
      MultipleSelection=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   795
      Left            =   14190
      MouseIcon       =   "frmCSMSEditCustomerVehicle_neo.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSEditCustomerVehicle_neo.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit Window"
      Top             =   7200
      Width           =   705
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   795
      Left            =   13500
      MouseIcon       =   "frmCSMSEditCustomerVehicle_neo.frx":153A
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSEditCustomerVehicle_neo.frx":168C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print this Record"
      Top             =   7200
      Width           =   705
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   14805
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   30
      Top             =   7110
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   795
      Left            =   12810
      MouseIcon       =   "frmCSMSEditCustomerVehicle_neo.frx":19F2
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSEditCustomerVehicle_neo.frx":1B44
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Add Vehicle"
      Top             =   7200
      Width           =   705
   End
   Begin VB.PictureBox picEDIT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5505
      Left            =   1470
      ScaleHeight     =   5475
      ScaleWidth      =   11805
      TabIndex        =   33
      Top             =   1290
      Visible         =   0   'False
      Width           =   11835
      Begin VB.ComboBox Cbo_bodytype 
         Height          =   330
         Left            =   5130
         TabIndex        =   94
         Top             =   4500
         Width           =   2355
      End
      Begin VB.Timer Timer1 
         Interval        =   400
         Left            =   8820
         Top             =   420
      End
      Begin MSComCtl2.DTPicker txtFDATE 
         Height          =   375
         Left            =   5820
         TabIndex        =   82
         Top             =   3300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   20250625
         CurrentDate     =   40091
      End
      Begin Crystal.CrystalReport rptVEHI 
         Left            =   7620
         Top             =   4140
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox txtTINNO 
         Height          =   345
         Left            =   1380
         TabIndex        =   11
         Top             =   3720
         Width           =   2325
      End
      Begin VB.ComboBox cboMODEL 
         Height          =   330
         Left            =   5130
         TabIndex        =   15
         Text            =   "cboMODEL"
         Top             =   4110
         Width           =   2355
      End
      Begin VB.TextBox txtIDATE_ 
         BackColor       =   &H000000FF&
         Height          =   345
         Left            =   10140
         TabIndex        =   28
         Top             =   4050
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox txtFDATE_ 
         BackColor       =   &H000000FF&
         Height          =   345
         Left            =   8520
         TabIndex        =   24
         Top             =   4050
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox txtDDATE_ 
         BackColor       =   &H000000FF&
         Height          =   345
         Left            =   10140
         TabIndex        =   20
         Top             =   3690
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox txtSDATE_ 
         BackColor       =   &H000000FF&
         Height          =   345
         Left            =   8520
         TabIndex        =   19
         Top             =   3690
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox txtWAR 
         Height          =   345
         Left            =   9420
         TabIndex        =   25
         Top             =   2070
         Width           =   2295
      End
      Begin VB.TextBox txtINV 
         Height          =   345
         Left            =   5820
         TabIndex        =   21
         Top             =   2070
         Width           =   1605
      End
      Begin VB.ComboBox cboCOLOR 
         Height          =   330
         Left            =   1380
         TabIndex        =   7
         Text            =   "cboCOLOR"
         Top             =   2100
         Width           =   2325
      End
      Begin VB.ComboBox cboEND 
         Height          =   330
         Left            =   5820
         TabIndex        =   18
         Text            =   "cboEND"
         Top             =   1260
         Width           =   5895
      End
      Begin VB.ComboBox cboSELLING 
         Height          =   330
         Left            =   5820
         TabIndex        =   17
         Text            =   "cboSELLING"
         Top             =   840
         Width           =   5895
      End
      Begin VB.TextBox txtICOMP 
         Height          =   345
         Left            =   9420
         TabIndex        =   27
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txtITYPE 
         Height          =   345
         Left            =   9420
         TabIndex        =   26
         Top             =   2490
         Width           =   2295
      End
      Begin VB.TextBox txtPROD 
         Height          =   345
         Left            =   1380
         TabIndex        =   9
         Top             =   2910
         Width           =   2295
      End
      Begin VB.TextBox txtFCOMP 
         Height          =   345
         Left            =   5820
         TabIndex        =   23
         Top             =   2910
         Width           =   2175
      End
      Begin VB.TextBox txtFTYPE 
         Height          =   345
         Left            =   5820
         TabIndex        =   22
         Top             =   2490
         Width           =   2175
      End
      Begin VB.ComboBox cboMCODE 
         Height          =   330
         Left            =   1380
         TabIndex        =   14
         Text            =   "cboMCODE"
         Top             =   4590
         Width           =   2355
      End
      Begin VB.ComboBox cboMAKE 
         Height          =   330
         Left            =   1380
         TabIndex        =   12
         Text            =   "cboMAKE"
         Top             =   4170
         Width           =   2355
      End
      Begin VB.ComboBox cboYEAR 
         Height          =   330
         Left            =   5130
         TabIndex        =   13
         Text            =   "cboYEAR"
         Top             =   3750
         Width           =   2355
      End
      Begin VB.TextBox txtSerialNo 
         Height          =   345
         Left            =   1380
         TabIndex        =   10
         Top             =   3330
         Width           =   2325
      End
      Begin VB.TextBox txtENGINE 
         Height          =   345
         Left            =   1380
         TabIndex        =   8
         Top             =   2490
         Width           =   2295
      End
      Begin VB.TextBox txtDESC 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1380
         TabIndex        =   16
         Top             =   4950
         Width           =   6105
      End
      Begin VB.TextBox txtVIN 
         Height          =   345
         Left            =   1380
         TabIndex        =   4
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtCONDNO 
         Height          =   345
         Left            =   1380
         TabIndex        =   6
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   915
         Left            =   10950
         MouseIcon       =   "frmCSMSEditCustomerVehicle_neo.frx":1E57
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSEditCustomerVehicle_neo.frx":1FA9
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Cancel"
         Top             =   4440
         Width           =   735
      End
      Begin VB.ComboBox cboNAME 
         Height          =   330
         Left            =   1380
         TabIndex        =   3
         Text            =   "cboNAME"
         Top             =   450
         Width           =   5355
      End
      Begin VB.TextBox txtPLATENO 
         Height          =   345
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   5
         Top             =   1260
         Width           =   2295
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   915
         Left            =   10230
         MouseIcon       =   "frmCSMSEditCustomerVehicle_neo.frx":22E7
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSEditCustomerVehicle_neo.frx":2439
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Save Entry"
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton cmdPrintVeh 
         Caption         =   "&Print"
         Height          =   915
         Left            =   9510
         MouseIcon       =   "frmCSMSEditCustomerVehicle_neo.frx":2789
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSEditCustomerVehicle_neo.frx":28DB
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Print this Record"
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   915
         Left            =   8790
         MouseIcon       =   "frmCSMSEditCustomerVehicle_neo.frx":2C41
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSEditCustomerVehicle_neo.frx":2D93
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Delete Selected Vehicle"
         Top             =   4440
         Width           =   735
      End
      Begin MSComCtl2.DTPicker txtIDATE 
         Height          =   375
         Left            =   9420
         TabIndex        =   83
         Top             =   3270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   20250625
         CurrentDate     =   40091
      End
      Begin MSComCtl2.DTPicker txtSDATE 
         Height          =   375
         Left            =   5820
         TabIndex        =   84
         Top             =   1650
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   20250625
         CurrentDate     =   40091
      End
      Begin MSComCtl2.DTPicker txtDDATE 
         Height          =   375
         Left            =   9420
         TabIndex        =   85
         Top             =   1650
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   20250625
         CurrentDate     =   40091
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BODY TYPE"
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
         Index           =   28
         Left            =   4170
         TabIndex        =   95
         Top             =   4560
         Width           =   885
      End
      Begin VB.Label lblloyal 
         AutoSize        =   -1  'True
         Caption         =   "** Loyalty Member **"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   10020
         TabIndex        =   86
         Top             =   450
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblINFO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIN NO"
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
         Index           =   25
         Left            =   780
         TabIndex        =   65
         Top             =   3840
         Width           =   525
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MODEL"
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
         Index           =   24
         Left            =   4530
         TabIndex        =   64
         Top             =   4200
         Width           =   570
      End
      Begin VB.Label LBLOLDVIN 
         BackColor       =   &H000000FF&
         Caption         =   "old vin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7620
         TabIndex        =   63
         Top             =   4980
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label LBLOLDPLATE 
         BackColor       =   &H000000FF&
         Caption         =   "old plate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7620
         TabIndex        =   62
         Top             =   4650
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INS EXP DATE"
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
         Index           =   23
         Left            =   8280
         TabIndex        =   58
         Top             =   3390
         Width           =   1065
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INS COMPANY"
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
         Index           =   22
         Left            =   8220
         TabIndex        =   57
         Top             =   3000
         Width           =   1125
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INS TYPE"
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
         Index           =   21
         Left            =   8640
         TabIndex        =   56
         Top             =   2610
         Width           =   705
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIN EXP DATE"
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
         Index           =   20
         Left            =   4695
         TabIndex        =   55
         Top             =   3420
         Width           =   1050
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIN COMPANY"
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
         Index           =   19
         Left            =   4635
         TabIndex        =   54
         Top             =   3000
         Width           =   1110
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIN TYPE"
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
         Index           =   18
         Left            =   5055
         TabIndex        =   53
         Top             =   2580
         Width           =   690
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "END USER"
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
         Index           =   17
         Left            =   4995
         TabIndex        =   52
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SELLING DEALER"
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
         Index           =   16
         Left            =   4410
         TabIndex        =   51
         Top             =   930
         Width           =   1335
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DELIVERY DATE"
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
         Index           =   15
         Left            =   8130
         TabIndex        =   50
         Top             =   1770
         Width           =   1230
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WARRANTY NO"
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
         Index           =   14
         Left            =   8160
         TabIndex        =   49
         Top             =   2190
         Width           =   1185
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE NO"
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
         Index           =   13
         Left            =   4830
         TabIndex        =   48
         Top             =   2190
         Width           =   915
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATE SOLD"
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
         Index           =   12
         Left            =   4845
         TabIndex        =   47
         Top             =   1770
         Width           =   900
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SERIAL NO"
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
         Index           =   11
         Left            =   465
         TabIndex        =   46
         Top             =   3450
         Width           =   840
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROD NO"
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
         Index           =   10
         Left            =   600
         TabIndex        =   45
         Top             =   3030
         Width           =   705
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENGINE NO"
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
         Index           =   9
         Left            =   480
         TabIndex        =   44
         Top             =   2580
         Width           =   825
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPTION"
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
         Index           =   8
         Left            =   255
         TabIndex        =   43
         Top             =   5040
         Width           =   1050
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MODEL CODE"
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
         Index           =   7
         Left            =   255
         TabIndex        =   42
         Top             =   4680
         Width           =   1050
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MAKE"
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
         Index           =   6
         Left            =   840
         TabIndex        =   41
         Top             =   4260
         Width           =   465
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "YEAR"
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
         Index           =   5
         Left            =   4650
         TabIndex        =   40
         Top             =   3840
         Width           =   420
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR"
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
         Index           =   4
         Left            =   735
         TabIndex        =   39
         Top             =   2190
         Width           =   570
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COND NO"
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
         Index           =   3
         Left            =   585
         TabIndex        =   38
         Top             =   1770
         Width           =   720
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLATE NO"
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
         Index           =   2
         Left            =   510
         TabIndex        =   37
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIN NO"
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
         Index           =   1
         Left            =   765
         TabIndex        =   36
         Top             =   930
         Width           =   540
      End
      Begin VB.Label lblINFO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OWNER NAME"
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
         Index           =   0
         Left            =   225
         TabIndex        =   35
         Top             =   540
         Width           =   1080
      End
      Begin VB.Label labid 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   10830
         TabIndex        =   34
         Top             =   60
         Visible         =   0   'False
         Width           =   870
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Index           =   0
         Left            =   -30
         TabIndex        =   81
         Top             =   0
         Width           =   11865
         _Version        =   655364
         _ExtentX        =   20929
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "Vehicle Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picDVIN 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3465
      Left            =   5520
      ScaleHeight     =   3435
      ScaleWidth      =   3855
      TabIndex        =   71
      Top             =   2340
      Visible         =   0   'False
      Width           =   3885
      Begin wizButton.cmd cmd1 
         Height          =   345
         Left            =   3510
         TabIndex        =   72
         Top             =   0
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         TX              =   "X"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCSMSEditCustomerVehicle_neo.frx":30BE
      End
      Begin MSComctlLib.ListView lsvDVIN 
         Height          =   2715
         Left            =   30
         TabIndex        =   73
         Top             =   420
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   4789
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "VIN NO"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "COUNT"
            Object.Width           =   1764
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Index           =   2
         Left            =   0
         TabIndex        =   75
         Top             =   0
         Width           =   11805
         _Version        =   655364
         _ExtentX        =   20823
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "DUPLICATE VIN NO."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
      End
      Begin VB.Label lblINFO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOUBLE CLICK TO DISPLAY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Index           =   27
         Left            =   1620
         TabIndex        =   74
         Top             =   3150
         Width           =   2175
      End
   End
   Begin VB.PictureBox picDPLATE 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3465
      Left            =   5850
      ScaleHeight     =   3435
      ScaleWidth      =   3195
      TabIndex        =   66
      Top             =   2318
      Visible         =   0   'False
      Width           =   3225
      Begin wizButton.cmd cmdx 
         Height          =   345
         Left            =   2850
         TabIndex        =   70
         Top             =   0
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         TX              =   "X"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCSMSEditCustomerVehicle_neo.frx":30DA
      End
      Begin MSComctlLib.ListView lsvDPLATE 
         Height          =   2715
         Left            =   30
         TabIndex        =   67
         Top             =   420
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4789
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PLATE NO"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "COUNT"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lblINFO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOUBLE CLICK TO DISPLAY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Index           =   26
         Left            =   975
         TabIndex        =   69
         Top             =   3150
         Width           =   2175
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Index           =   1
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Width           =   11805
         _Version        =   655364
         _ExtentX        =   20823
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "DUPLICATE PLATE NO."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
      End
   End
   Begin VB.PictureBox PICLoyaltyID 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1785
      Left            =   5715
      ScaleHeight     =   1755
      ScaleWidth      =   3465
      TabIndex        =   89
      Top             =   3113
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox txtLoyalID 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   60
         TabIndex        =   91
         ToolTipText     =   "Type Loyalty ID here"
         Top             =   390
         Width           =   3345
      End
      Begin VB.CommandButton CmdloyaltyCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   2700
         MouseIcon       =   "frmCSMSEditCustomerVehicle_neo.frx":30F6
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSEditCustomerVehicle_neo.frx":3248
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Cancel"
         Top             =   840
         Width           =   705
      End
      Begin VB.CommandButton CmdloyaltySave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   2010
         MouseIcon       =   "frmCSMSEditCustomerVehicle_neo.frx":3586
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSEditCustomerVehicle_neo.frx":36D8
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Save this Record"
         Top             =   840
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   315
         Left            =   0
         TabIndex        =   93
         Top             =   -30
         Width           =   4125
         _Version        =   655364
         _ExtentX        =   7276
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Loyalty Identification no."
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   12582912
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F7 - INPUT LOYALTY ID"
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
      Left            =   9030
      TabIndex        =   88
      Top             =   7740
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F3 - TO SEARCH"
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
      Left            =   150
      TabIndex        =   87
      Top             =   7740
      Width           =   1245
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   79
      Top             =   0
      Width           =   14955
      _Version        =   655364
      _ExtentX        =   26379
      _ExtentY        =   556
      _StockProps     =   14
      Caption         =   "Search (Type your keyword here)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   77
      Top             =   270
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F6 - DISPLAY DUPLICATE VIN NO."
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
      Left            =   6210
      TabIndex        =   61
      Top             =   7740
      Width           =   2595
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F5 - REFRESH"
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
      Left            =   4830
      TabIndex        =   60
      Top             =   7740
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F4 - DISPLAY DUPLICATE PLATE NO."
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
      Index           =   0
      Left            =   1650
      TabIndex        =   59
      Top             =   7740
      Width           =   2850
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   78
      Top             =   7680
      Width           =   12705
      _Version        =   655364
      _ExtentX        =   22410
      _ExtentY        =   556
      _StockProps     =   14
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      ForeColor       =   4194304
   End
End
Attribute VB_Name = "frmCSMSEditCustomerVehicle_neo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VSEARCHMODE                                        As String
Dim AddorEdit                                          As String

Private Sub CmdloyaltyCancel_Click()
    PICLoyaltyID.Visible = False
    picEDIT.Enabled = True
    
    Call cmdCancel_Click
End Sub

Private Sub CmdloyaltySave_Click()
    Dim rsLoyaltyID                         As New ADODB.Recordset
    Dim LID                                 As String
    Dim CTR                                 As Long
    
    On Error GoTo Error
    
    If MsgBox("Register this loyalty id to this Vehicle, Are you sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    LID = N2Str2Null(txtLoyalID.Text)
    Set rsLoyaltyID = gconDMIS.Execute("select id, loyalty_id from CSMS_CUSVEH where loyalty_id = " & LID & "")
    If Not (rsLoyaltyID.BOF And rsLoyaltyID.EOF) Then
        If rsLoyaltyID!ID <> labid Then
            MessagePop RecSaveWarning, "Duplicate Record", "Loyalty ID Already registered to another Vehicle"
            Exit Sub
        End If
    End If

    gconDMIS.Execute ("Update CSMS_CUSVEH set " & _
        " Loyalty_ID = " & LID & _
        " where id = " & labid.Caption & "")
        
    MessagePop RecSaveOk, "Info", "Record Saved"
    
    Call CmdloyaltyCancel_Click
    Exit Sub
    
Error:
    Call ErrHandler(gconDMIS)
End Sub


Sub ErrHandler(objCon As Object)
    Dim ADOErr As ADODB.Error
    Dim strError As String
    
    For Each ADOErr In objCon.Errors
        strError = "Error #: " & ADOErr.Number & vbCrLf & _
            "Error Description : " & ADOErr.Description
        
    Next
    
    MsgBox strError, vbCritical, "Error"
    objCon.Errors.Clear
End Sub

Function FindModel(VMODEL As String)
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("SELECT * FROM ALL_MODEL WHERE LTRIM(RTRIM(CODE)) = '" & VMODEL & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        FindModel = LTrim(RTrim(Null2String(rstmp!DESCRIPT)))
    Else
        FindModel = ""
    End If
    Set rstmp = Nothing
End Function

Function FindCustomerCode(VNAME As String) As String
    Dim rsKUTO                                         As New ADODB.Recordset
    Set rsKUTO = gconDMIS.Execute("SELECT * FROM ALL_CUSTOMER_TABLE WHERE ACCTNAME = '" & VNAME & "'")
    If Not (rsKUTO.BOF And rsKUTO.EOF) Then
        FindCustomerCode = RTrim(LTrim(rsKUTO!CUSCDE))
    Else
        FindCustomerCode = ""
    End If
    Set rsKUTO = Nothing
End Function

Function SetColor(CCC As String)
    Dim rsColor                                        As ADODB.Recordset
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select COLOR_CODE,COLOR_DESC from ALL_Color where COLOR_DESC = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        SetColor = Null2String(rsColor!Color_code)
    Else
        SetColor = ""
    End If
End Function

Function SetEndUser(XXX As String, CodeOrName As Integer) As String
    Dim rsEndUser                                      As ADODB.Recordset
    Set rsEndUser = New ADODB.Recordset
    Dim SelectionCodeOrName                            As String
    If CodeOrName = 1 Then
        SelectionCodeOrName = "CusCde"
    Else
        SelectionCodeOrName = "AcctName"
    End If
    Set rsEndUser = gconDMIS.Execute("Select * from All_Customer Where " & SelectionCodeOrName & " = '" & XXX & "'")
    If Not rsEndUser.EOF And Not rsEndUser.BOF Then
        If CodeOrName = 1 Then
            SetEndUser = Null2String(rsEndUser!AcctName)
        Else
            SetEndUser = Null2String(rsEndUser!CUSCDE)
        End If
    End If
End Function

Function SetSellingDealer(XXX As String, CodeOrName As Integer) As String
    Dim rsSellingDealer                                As ADODB.Recordset
    Set rsSellingDealer = New ADODB.Recordset
    Dim SelectionCodeOrName                            As String
    If CodeOrName = 1 Then
        SelectionCodeOrName = "DealerCode"
    Else
        SelectionCodeOrName = "DealerName"
    End If
    Set rsSellingDealer = gconDMIS.Execute("Select * from CSMS_SellingDealer Where " & SelectionCodeOrName & " = '" & XXX & "'")
    If Not rsSellingDealer.EOF And Not rsSellingDealer.BOF Then
        If CodeOrName = 1 Then
            SetSellingDealer = Null2String(rsSellingDealer!dealername)
        Else
            SetSellingDealer = Null2String(rsSellingDealer!DEALERCODE)
        End If
    End If
End Function

Function ReturnVehicleID(VKEY As String, vFIELD As String, VEHID As Integer, KKEY As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("select * from csms_Cusveh where " & vFIELD & " = '" & VKEY & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        If VEHID = rstmp!ID Then
            ReturnVehicleID = False
        Else
            MsgBox "" & KKEY & " already exist, Registered to" & vbCrLf & GetAcctName(Null2String(rstmp!CUSCDE)) & "", vbExclamation, "CSMS"
            ReturnVehicleID = True
        End If
    End If
    Set rstmp = Nothing
End Function

Function GetAcctName(ACCTNO As String) As String
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE CUSCDE = '" & ACCTNO & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        GetAcctName = Null2String(rstmp!AcctName)
    End If
    Set rstmp = Nothing
End Function

Function CheckIfPlateNoAlreadyExist(vFIELD As String, VKEY As String, VTYPE As Integer, VFNAME As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("Select " & vFIELD & ",NIYM From CSMS_CusVeh Where " & vFIELD & " = '" & VKEY & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        If VTYPE = 1 Then MsgBox "" & VFNAME & " " & VKEY & " already exist and registered to " & rstmp!NIYM & "", vbInformation, "CSMS"
        CheckIfPlateNoAlreadyExist = True
    Else
        CheckIfPlateNoAlreadyExist = False
    End If

    Set rstmp = Nothing
End Function

Function CheckIfVinNoAlreadyExist(vVin As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("SELECT ID FROM CSMS_CUSVEH WHERE VIN = '" & vVin & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        CheckIfVinNoAlreadyExist = True
    Else
        CheckIfVinNoAlreadyExist = False
    End If
    Set rstmp = Nothing
End Function

Function FindVehicleID(VPLATE As String, vFIELD As String, tmp As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_CUSVEH WHERE " & vFIELD & " = '" & VPLATE & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        If Not labid.Caption = rstmp!ID Then
            MsgBox "" & tmp & " no. " & VPLATE & " is already registered to " & LTrim(RTrim(rstmp!NIYM)) & "", vbInformation, "Vehicle Already Exist"
            FindVehicleID = False
        Else
            FindVehicleID = True
        End If
    End If

    Set rstmp = Nothing
End Function

Function FindColorDESC(CCODE As String)
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT * FROM ALL_COLOR WHERE COLOR_CODE = '" & CCODE & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        FindColorDESC = LTrim(RTrim(Null2String(rstmp!color_desc)))
    Else
        FindColorDESC = ""
    End If
    Set rstmp = Nothing
End Function

Function FindEndUser(ECODE As String)
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT * FROM ALL_CUSTOMER_TABLE WHERE CUSCDE = '" & ECODE & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        FindEndUser = Null2String(rstmp!AcctName)
    Else
        FindEndUser = ""
    End If
    Set rstmp = Nothing
End Function

Function FindSDName(DCODE As String)
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_SELLINGDEALER WHERE DEALERCODE = '" & DCODE & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        FindSDName = Null2String(rstmp!dealername)
    Else
        FindSDName = ""
    End If
    Set rstmp = Nothing
End Function

Public Function flex_FillReportView(RS As ADODB.Recordset, grd As XtremeReportControl.ReportControl, Optional ByVal WithSN As Boolean = False)
    Dim fld                                            As ADODB.FIELD
    Dim j                                              As Long
    Dim REC                                            As XtremeReportControl.ReportRecord

    grd.Records.DeleteAll

    While Not RS.EOF
        j = j + 1

        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem j
        End If
        For Each fld In RS.Fields
            REC.AddItem (Trim(fld.Value))
        Next
        RS.MoveNext
    Wend
    grd.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RS = Nothing
End Function

Sub initMemvars()
    cboNAME.ListIndex = 0
    txtVIN.Text = ""
    txtPLATENO.Text = ""
    txtCONDNO.Text = ""
    cboCOLOR.ListIndex = 0
    txtENGINE.Text = ""
    txtPROD.Text = ""
    txtSerialNo.Text = ""
    txtTINNO.Text = ""
    cboMAKE.ListIndex = 0
    cboYear.ListIndex = 0
    cboMCODE.ListIndex = 0
    cboModel.ListIndex = 0
    txtDesc.Text = ""
    cboSELLING.ListIndex = 0
    cboEND.ListIndex = 0
    'txtSDATE.Text = ""
    txtSDATE.CheckBox = False
    txtINV.Text = ""
    txtFTYPE.Text = ""
    txtFCOMP.Text = ""
    txtFDATE.CheckBox = False
    'txtDDATE.Text = ""
    txtDDATE.CheckBox = False
    txtWAR.Text = ""
    txtITYPE.Text = ""
    txtICOMP.Text = ""
    txtIDATE.CheckBox = False
    Cbo_bodytype.Text = ""
End Sub

Sub DisplayDuplicateVIN()
    Dim rstmp                                          As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    lsvDVIN.ListItems.Clear
    Set rstmp = gconDMIS.Execute("SELECT VIN,COUNT(VIN) AS COUNTS FROM CSMS_CUSVEH GROUP BY VIN HAVING COUNT(VIN) > 1 ORDER BY VIN")
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            Set ITEM = lsvDVIN.ListItems.Add(, , Null2String(rstmp!Vin))
            ITEM.SubItems(1) = rstmp!COUNTS
            rstmp.MoveNext
        Loop
    End If

    Set rstmp = Nothing
End Sub

Sub DisplayDuplicatePlate()
    Dim rstmp                                          As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    lsvDPLATE.ListItems.Clear
    Set rstmp = gconDMIS.Execute("SELECT PLATE_NO,COUNT(PLATE_NO) AS COUNTS FROM CSMS_CUSVEH GROUP BY PLATE_NO HAVING COUNT(PLATE_NO) > 1 ORDER BY PLATE_NO")
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            Set ITEM = lsvDPLATE.ListItems.Add(, , Null2String(rstmp!PLATE_NO))
            ITEM.SubItems(1) = rstmp!COUNTS
            rstmp.MoveNext
        Loop
    End If

    Set rstmp = Nothing
End Sub

Sub FillMOdelCode()
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT DISTINCT(MODELCODE) FROM CSMS_CUSVEH ORDER BY MODELCODE")
    cboMCODE.Clear
    If Not (rstmp.EOF And rstmp.BOF) Then
        Do While Not rstmp.EOF
            cboMCODE.AddItem LTrim(RTrim(Null2String(rstmp!MODELCODE)))
            rstmp.MoveNext
        Loop
    End If
    Set rstmp = Nothing
End Sub

Sub FillModel()
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT DISTINCT(MODEL) FROM CSMS_CUSVEH ORDER BY MODEL")
    cboModel.Clear
    If Not (rstmp.EOF And rstmp.BOF) Then
        Do While Not rstmp.EOF
            cboModel.AddItem LTrim(RTrim(Null2String(rstmp!Model)))
            rstmp.MoveNext
        Loop
    End If
    Set rstmp = Nothing
End Sub

Sub FillSellingDealer()
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_SELLINGDEALER ORDER BY DEALERNAME")
    cboSELLING.Clear
    If Not (rstmp.EOF And rstmp.BOF) Then
        Do While Not rstmp.EOF
            cboSELLING.AddItem Null2String(rstmp!dealername)
            rstmp.MoveNext
        Loop
    End If
    Set rstmp = Nothing
End Sub

Sub FillCustomer()
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT * FROM all_CUSTOMER_TABLE ORDER BY ACCTNAME")
    cboNAME.Clear
    cboEND.Clear
    If Not (rstmp.EOF And rstmp.BOF) Then
        Do While Not rstmp.EOF
            cboNAME.AddItem Null2String(rstmp!AcctName)
            cboEND.AddItem Null2String(rstmp!AcctName)
            rstmp.MoveNext
        Loop
    End If
    Set rstmp = Nothing
End Sub

Sub FillMake()
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT DISTINCT(MAKE) FROM ALL_MAKE ORDER BY MAKE")
    cboMAKE.Clear
    If Not (rstmp.EOF And rstmp.BOF) Then
        Do While Not rstmp.EOF
            cboMAKE.AddItem LTrim(RTrim(Null2String(rstmp!Make)))
            rstmp.MoveNext
        Loop
    End If
    Set rstmp = Nothing
End Sub

Sub FillColorS()
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT * FROM all_COLOR ORDER BY COLOR_DESC")
    cboCOLOR.Clear
    If Not (rstmp.EOF And rstmp.BOF) Then
        Do While Not rstmp.EOF
            cboCOLOR.AddItem LTrim(RTrim(Null2String(rstmp!color_desc)))
            rstmp.MoveNext
        Loop
    End If
    Set rstmp = Nothing
End Sub

Sub displayDuplicateVinNo(vVin As String)
    Screen.MousePointer = 11
    Call ReportControlAddColumnHeader(rptVEH, " , NAME, PLATE NO., COND. NO., VIN NO., DESCRIPTION, ENGINE NO., DATE SOLD, DATE DELIVERY, , MAKE")
    Call ReportControlPaintManager(rptVEH)
    rptVEH.GroupsOrder.Add rptVEH.Columns(0)
    rptVEH.Columns(0).Visible = False
    rptVEH.Columns(9).Visible = False
    Call ResizeColumnHeader(rptVEH, "0, 30, 11, 11, 19, 29, 15, 12, 16, 0, 10")
    Call flex_FillReportView(gconDMIS.Execute("SELECT CUSCDE, NIYM, PLATE_NO, VCOND_NO, VIN, DESCRIPTION, ENGINE, D_SOLD, DEL_DATE, ID, MAKE FROM CSMS_CUSVEH WHERE VIN = '" & vVin & "'"), rptVEH)

    Screen.MousePointer = 0
End Sub

Sub displayDuplicatePlateNo(VPLATE As String)
    Screen.MousePointer = 11
    Call ReportControlAddColumnHeader(rptVEH, " , NAME, PLATE NO., COND. NO., VIN NO., DESCRIPTION, ENGINE NO., DATE SOLD, DATE DELIVERY, , MAKE")
    Call ReportControlPaintManager(rptVEH)
    rptVEH.GroupsOrder.Add rptVEH.Columns(0)
    rptVEH.Columns(0).Visible = False
    rptVEH.Columns(9).Visible = False
    Call ResizeColumnHeader(rptVEH, "0, 30, 11, 11, 19, 29, 15, 12, 16, 0, 10")
    Call flex_FillReportView(gconDMIS.Execute("SELECT CUSCDE, NIYM, PLATE_NO, VCOND_NO, VIN, DESCRIPTION, ENGINE, D_SOLD, DEL_DATE, ID, MAKE FROM CSMS_CUSVEH WHERE PLATE_NO = '" & VPLATE & "'"), rptVEH)

    Screen.MousePointer = 0
End Sub

Sub displayVehicle()
    Screen.MousePointer = 11
    Call ReportControlAddColumnHeader(rptVEH, " , NAME, PLATE NO., COND. NO., VIN NO., DESCRIPTION, ENGINE NO., DATE SOLD, DATE DELIVERY, , MAKE")
    Call ReportControlPaintManager(rptVEH)
    rptVEH.GroupsOrder.Add rptVEH.Columns(0)
    rptVEH.Columns(0).Visible = False
    rptVEH.Columns(9).Visible = False
    Call ResizeColumnHeader(rptVEH, "0, 30, 11, 11, 19, 29, 15, 12, 16, 0, 10")
    Call flex_FillReportView(gconDMIS.Execute("SELECT  CUSCDE, NIYM, PLATE_NO, VCOND_NO, VIN, DESCRIPTION, ENGINE, D_SOLD, DEL_DATE, ID, MAKE FROM CSMS_CUSVEH"), rptVEH)

    Screen.MousePointer = 0
End Sub

Sub ReportControlAddColumnHeader(LST As ReportControl, StringHeaders As String)
    Dim ar()                                           As String
    Dim i                                              As Integer

    ar = Split(StringHeaders, ",")
    LST.Columns.DeleteAll
    For i = LBound(ar) To UBound(ar)
        LST.Columns.Add i, ar(i), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Sub ReportControlPaintManager(LST As ReportControl)
    With LST
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.ColumnStyle = xtpColumnExplorer
    End With
End Sub

Sub FillOtherInfo(vID As Integer)
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_CUSVEH WHERE ID = " & labid & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        cboCOLOR.Text = FindColorDESC(LTrim(RTrim(Null2String(rstmp!ClrCde))))
        txtPROD.Text = Null2String(rstmp!prodno)
        txtSerialNo.Text = Null2String(rstmp!SERIAL)
        txtTINNO.Text = Null2String(rstmp!TIN_Number)
        cboYear.Text = Null2String(rstmp!YER)
        cboMAKE.Text = Null2String(rstmp!Make)
        cboMCODE.Text = LTrim(RTrim(Null2String(rstmp!MODELCODE)))
        cboModel.Text = Null2String(rstmp!Model)
        txtDesc.Text = Null2String(rstmp!Description)
        cboSELLING.Text = FindSDName(Null2String(rstmp!Selling_Dealer))
        cboEND.Text = FindEndUser(Null2String(rstmp!ENDUSER))
        'txtSDATE.Text = Null2String(RSTMP!D_SOLD)
        txtSDATE.Value = Null2String(rstmp!D_SOLD)
        txtINV.Text = Null2String(rstmp!INVOICENO)
        'txtDDATE.Text = Null2String(RSTMP!DEL_DATE)
        txtDDATE.Value = Null2String(rstmp!DEL_DATE)
        txtWAR.Text = Null2String(rstmp!War_Cert)

        txtFTYPE.Text = Null2String(rstmp!FIN_TYPE)
        txtFCOMP.Text = Null2String(rstmp!FIN_COMP)
        'txtFDATE.Text = Null2Date(RSTMP!FIN_EXP_DATE)
        txtFDATE.Value = Null2Date(rstmp!FIN_EXP_DATE)

        txtITYPE.Text = Null2String(rstmp!INS_TYPE)
        txtICOMP.Text = Null2String(rstmp!INS_COMP)
        'txtIDATE.Text = Null2String(RSTMP!INS_EXP_DATE)
        txtIDATE.Value = Null2String(rstmp!INS_EXP_DATE)
        
        
        If COMPANY_CODE = "HAI" Then
            txtLoyalID = Null2String(rstmp!Loyalty_ID)
            
            If (rstmp!Loyalty_ID) <> "" Then
               lblloyal.Visible = True
            Else
               lblloyal.Visible = False
            End If
        End If
    End If
    Set rstmp = Nothing
End Sub


Private Sub cboMAKE_GotFocus()
    'UPDATE BY: JUN 10152008
    Call FillMake
End Sub

Private Sub cboMCODE_Change()
    'txtModel.Text = FindModel(cboMCODE)
    'UPDATED BY: 10152008
    Call FillMOdelCode
End Sub

Private Sub cmd1_Click()
    picDVIN.Visible = False
    picDVIN.ZOrder 1

    txtSearch.Enabled = True
    rptVEH.Enabled = True
    cmdPrint.Enabled = True
    cmdExit.Enabled = True

    VSEARCHMODE = "NO"
    txtSearch.SetFocus
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "CUSTOMER VEHICLE") = False Then Exit Sub
    cmdDelete.Visible = False
    VSEARCHMODE = "YES"
    AddorEdit = "ADD"

    Call initMemvars

    picEDIT.Visible = True
    picEDIT.ZOrder 0

    txtSearch.Enabled = False
    rptVEH.Enabled = False
    cmdPrint.Enabled = False
    cmdExit.Enabled = False
End Sub

Private Sub cmdCancel_Click()

    picEDIT.Visible = False
    picEDIT.ZOrder 1

    txtSearch.Enabled = True
    rptVEH.Enabled = True
    cmdPrint.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "CUSTOMER VEHICLE") = False Then Exit Sub

    If MsgBox("Delete this Vehicle Information", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

    SQL_STATEMENT = "DELETE FROM CSMS_CUSVEH WHERE ID = " & labid & ""
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("X", "CUSTOMER VEHICLE", SQL_STATEMENT, labid, "", "PLATE NO: " & txtPLATENO, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    ShowDeletedMsg

    Call cmdCancel_Click
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "CUSTOMER VEHICLE") = False Then Exit Sub
    CrystalReport1.Formulas(0) = "companyname='" & COMPANY_NAME & "'"
    CrystalReport1.Formulas(1) = "COMPANYADDRESS='" & COMPANY_ADDRESS & "'"
    CrystalReport1.WindowTitle = "Vehicle Master File"
    PrintSQLReport CrystalReport1, CSMS_REPORT_PATH & "cusvehreports.rpt", "", DMIS_REPORT_Connection, 1
End Sub

Private Sub cmdPrintVeh_Click()
    If Function_Access(LOGID, "Acess_PRINT", "CUSTOMER VEHICLE") = False Then Exit Sub
    rptVEHI.Formulas(0) = "companyname='" & COMPANY_NAME & "'"
    rptVEHI.Formulas(1) = "COMPANYADDRESS='" & COMPANY_ADDRESS & "'"
    rptVEHI.WindowTitle = "Vehicle Information"
    PrintSQLReport rptVEHI, CSMS_REPORT_PATH & "Vehicle.rpt", "{CSMS_CUSVEH.PLATE_NO} = '" & txtPLATENO.Text & "'", DMIS_REPORT_Connection, 1

    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("V", "CUSTOMER VEHICLE", "", labid, "", "PLATE NO: " & txtPLATENO, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
End Sub

Private Sub cmdSave_Click()
    Dim TMP_PLATE                                      As String
    Dim TMP_VIN                                        As String

    Dim vCUSTNAME                                      As String
    Dim vCUSCDE                                        As String
    Dim VVINNO                                         As String
    Dim vPLATENO                                       As String
    Dim VCONDNO                                        As String
    Dim vCOLOR                                         As String
    Dim VENGINENO                                      As String
    Dim VPRODNO                                        As String
    Dim vSERIALNO                                      As String
    Dim vTINNO                                         As String
    Dim vWARCERT                                       As String
    Dim VYEAR                                          As String
    Dim VMAKE                                          As String
    Dim VMODELCODE                                     As String
    Dim VMODEL                                         As String
    Dim VDESCRIPTION                                   As String
    Dim VSELDEL                                        As String
    Dim vEndUser                                       As String
    Dim VDSOLD                                         As String
    Dim VINV                                           As String
    Dim vFINTYPE                                       As String
    Dim VFINNAME                                       As String
    Dim VFEXPDATE                                      As String
    Dim VDDATE                                         As String
    Dim VWAR                                           As String
    Dim vINSTYPE                                       As String
    Dim vINSNAME                                       As String
    Dim VIEXPDATE                                      As String


    If LTrim(RTrim(cboNAME.Text)) = "" Then
        ShowIsRequiredMsg ("Name Cannot be Blank")
        cboNAME.SetFocus
        Exit Sub
    End If

    If txtPLATENO.Text = "" Then
        ShowIsRequiredMsg ("Plate no Cannot be Blank")
        txtPLATENO.SetFocus
        Exit Sub
    End If

    If RTrim(LTrim(cboMAKE.Text)) = "" Then
        ShowIsRequiredMsg "Make cannot be blank"
        cboMAKE.SetFocus
        Exit Sub
    End If

    If RTrim(LTrim(cboYear.Text)) = "" Then
        ShowIsRequiredMsg "Year cannot be blank"
        cboYear.SetFocus
        Exit Sub
    End If

    '****************************************************************************************

    TMP_PLATE = txtPLATENO.Text
    TMP_VIN = txtVIN.Text

    If AddorEdit = "ADD" Then
        If CheckIfPlateNoAlreadyExist("PLATE_NO", TMP_PLATE, 1, "Plate no.") = True Then
            txtPLATENO.SetFocus
            Exit Sub
        End If
    Else
        If CheckIfPlateNoAlreadyExist("PLATE_NO", TMP_PLATE, 0, "Plate no.") = True Then
            If ReturnVehicleID(TMP_PLATE, "PLATE_NO", labid.Caption, "Plate no.") = True Then
                txtPLATENO.SetFocus
                Exit Sub
            End If
        End If
    End If

    If AddorEdit = "YES" Then
        If Not txtVIN.Text = "" Then
            If CheckIfPlateNoAlreadyExist("VIN", TMP_VIN, 1, "Vin no.") Then
                txtVIN.SetFocus
                Exit Sub
            End If
        End If
    Else
        If Not txtVIN.Text = "" Then
            If CheckIfPlateNoAlreadyExist("VIN", TMP_VIN, 0, "Vin no.") Then
                If ReturnVehicleID(TMP_VIN, "VIN", labid, "Vin no") = True Then
                    txtVIN.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

    vCUSCDE = FindCustomerCode(cboNAME)
    If vCUSCDE = "" Then
        MsgBox "Customer name not found in the Customer master file", vbInformation, "CSMS"
        cboNAME.SetFocus
        Exit Sub
    End If
    vCUSCDE = N2Str2Null(vCUSCDE)

    vEndUser = FindCustomerCode(cboEND)
    If Not cboEND.Text = "" Then
        If vEndUser = "" Then
            MsgBox "End user name not found in the Customer master file", vbInformation, "CSMS"
            cboEND.SetFocus
            Exit Sub
        End If
    End If

    vEndUser = N2Str2Null(vEndUser)
    vCUSTNAME = N2Str2Null(cboNAME)
    VVINNO = N2Str2Null(txtVIN)
    vPLATENO = N2Str2Null(txtPLATENO)
    VCONDNO = N2Str2Null(txtCONDNO)
    vCOLOR = N2Str2Null(SetColor(cboCOLOR))
    VENGINENO = N2Str2Null(txtENGINE)
    VPRODNO = N2Str2Null(txtPROD)
    vSERIALNO = N2Str2Null(txtSerialNo)
    vTINNO = N2Str2Null(txtTINNO)
    vWARCERT = N2Str2Null(txtWAR)
    VYEAR = N2Str2Null(cboYear)
    VMAKE = N2Str2Null(cboMAKE)
    VMODELCODE = N2Str2Null(cboMCODE)
    VMODEL = N2Str2Null(cboModel)
    VDESCRIPTION = N2Str2Null(txtDesc)
    VDSOLD = N2Str2Null(txtSDATE)
    VINV = N2Str2Null(txtINV)
    vFINTYPE = N2Str2Null(txtFTYPE)
    VFINNAME = N2Str2Null(txtFCOMP)
    VFEXPDATE = N2Str2Null(txtFDATE)
    VDDATE = N2Str2Null(txtDDATE)
    VWAR = N2Str2Null(txtWAR)
    vINSTYPE = N2Str2Null(txtITYPE)
    vINSNAME = N2Str2Null(txtICOMP)
    VIEXPDATE = N2Str2Null(txtIDATE)

    Dim SQL                                            As String
    Dim rsREPOR                                        As New ADODB.Recordset
    
    On Error GoTo ERROR_MSG
    gconDMIS.BeginTrans
    If AddorEdit = "ADD" Then
        SQL = "Insert into CSMS_Cusveh "
        SQL = SQL & " (CUSCDE, NIYM, Yer, Make, MODEL, ENGINE, PLATE_NO, CLRCDE, SERIAL, PRODNO, TIN_NUMBER, WAR_CERT, VIN, VCOND_NO, D_SOLD, Description, DEL_DATE, SELLING_DEALER, ENDUSER, FIN_TYPE, FIN_COMP, FIN_EXP_DATE, INS_TYPE, INS_COMP, INS_EXP_DATE) VALUES("
        SQL = SQL & vCUSCDE & ","
        SQL = SQL & vCUSTNAME & ","
        SQL = SQL & VYEAR & ","
        SQL = SQL & VMAKE & ","
        SQL = SQL & VMODEL & ","
        SQL = SQL & VENGINENO & ","
        SQL = SQL & vPLATENO & ","
        SQL = SQL & vCOLOR & ","
        SQL = SQL & vSERIALNO & ","
        SQL = SQL & VPRODNO & ","
        SQL = SQL & vTINNO & ","
        SQL = SQL & vWARCERT & ","
        SQL = SQL & VVINNO & ","
        SQL = SQL & VCONDNO & ","
        SQL = SQL & VDSOLD & ","
        SQL = SQL & VDESCRIPTION & ","
        SQL = SQL & VDDATE & ","
        SQL = SQL & N2Str2Null(SetSellingDealer(cboSELLING.Text, 2)) & ","
        SQL = SQL & N2Str2Null(SetEndUser(cboEND.Text, 2)) & ","
        SQL = SQL & vFINTYPE & ","
        SQL = SQL & VFINNAME & ","
        SQL = SQL & VFEXPDATE & ","
        SQL = SQL & vINSTYPE & ","
        SQL = SQL & vINSNAME & ","
        SQL = SQL & VIEXPDATE & ")"


        gconDMIS.Execute (SQL)
        SQL_STATEMENT = SQL
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("A", "CUSTOMER VEHICLE", SQL_STATEMENT, FindTransactionID(Null2String(vPLATENO), "PLATE_NO", "CSMS_CUSVEH"), "", "PLATE NO: " & Null2String(vPLATENO), "", "")
        'NEW LOG AUDIT-----------------------------------------------------
                        
                        
        Call ShowSuccessFullyAdded
    Else
        Set rsREPOR = gconDMIS.Execute("SELECT PLATE_NO FROM CSMS_REPOR WHERE PLATE_NO = '" & LBLOLDPLATE.Caption & "'")
        If Not (rsREPOR.BOF And rsREPOR.EOF) Then
            If MsgBox("This Vehicle had a Previous Transaction, Editing this vehicle will update all past transaction", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

            SQL_STATEMENT = "UPDATE CSMS_REPOR SET PLATE_NO = " & vPLATENO & _
                ", VIN = " & VVINNO & _
                ", MODEL = " & VMODEL & _
                " WHERE PLATE_NO = '" & LBLOLDPLATE.Caption & "'"
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT-----------------------------------------------------
                Dim rstmp                                  As New ADODB.Recordset
                Set rstmp = gconDMIS.Execute("SELECT REP_OR FROM CSMS_REPOR WHERE PLATE_NO = " & vPLATENO & "")
                If Not (rstmp.BOF And rstmp.EOF) Then
                    Do While Not rstmp.EOF
                        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(Null2String(rstmp!REP_OR)), "REP_OR", "CSMS_REPOR"), "", "PLATE NO: " & Null2String(vPLATENO), "", "")
    
                        rstmp.MoveNext
                    Loop
                End If
                Set rstmp = Nothing
            'NEW LOG AUDIT-----------------------------------------------------

            gconDMIS.Execute ("UPDATE CSMS_REPAIRORDER SET PLATE_NO = " & vPLATENO & ",MODEL = " & VMODEL & " WHERE PLATE_NO = '" & LBLOLDPLATE.Caption & "'")
            SQL_STATEMENT = "UPDATE CSMS_APPOINTMENT SET PLATE_NO = " & vPLATENO & ",MODEL = " & VMODEL & " WHERE PLATE_NO = '" & LBLOLDPLATE.Caption & "'"
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT-----------------------------------------------------
                Set rstmp = gconDMIS.Execute("SELECT APPTNO FROM CSMS_APPOINTMENT WHERE PLATE_NO = " & vPLATENO & "")
                If Not (rstmp.BOF And rstmp.EOF) Then
                    Do While Not rstmp.EOF
                        Call NEW_LogAudit("E", "APPOINTMENT", SQL_STATEMENT, FindTransactionID(N2Str2Null(Null2String(rstmp!APPTNO)), "APPTNO", "CSMS_APPOINTMENT"), "", "PLATE NO: " & Null2String(vPLATENO), "", "")
                        rstmp.MoveNext
                    Loop
                End If
                Set rstmp = Nothing
            'NEW LOG AUDIT-----------------------------------------------------
            
            SQL_STATEMENT = "UPDATE CSMS_ESTHD SET PLATE_NO = " & vPLATENO & _
                ", VIN = " & VVINNO & _
                ", MODEL = " & VMODEL & _
                " WHERE PLATE_NO = '" & LBLOLDPLATE.Caption & "'"
            gconDMIS.Execute (SQL_STATEMENT)
        End If

        SQL = "UPDATE CSMS_Cusveh SET "
        SQL = SQL & " CUSCDE = " & vCUSCDE & ","
        SQL = SQL & " NIYM = " & vCUSTNAME & ","
        SQL = SQL & " Yer = " & VYEAR & ","
        SQL = SQL & " Make = " & VMAKE & ","
        SQL = SQL & " MODELCODE = " & VMODELCODE & ","
        SQL = SQL & " MODEL = " & VMODEL & ","
        SQL = SQL & " ENGINE = " & VENGINENO & ","
        SQL = SQL & " PLATE_NO = " & vPLATENO & ","
        SQL = SQL & " CLRCDE = " & vCOLOR & ","
        SQL = SQL & " SERIAL = " & vSERIALNO & ","
        SQL = SQL & " PRODNO = " & VPRODNO & ","
        SQL = SQL & " TIN_NUMBER = " & vTINNO & ","
        SQL = SQL & " WAR_CERT = " & vWARCERT & ","
        SQL = SQL & " VIN = " & VVINNO & ","
        SQL = SQL & " VCOND_NO = " & VCONDNO & ","
        SQL = SQL & " D_SOLD =" & VDSOLD & ","
        SQL = SQL & " DEL_DATE =" & VDDATE & ","
        SQL = SQL & " Description = " & VDESCRIPTION & ","
        SQL = SQL & " Selling_Dealer = " & N2Str2Null(SetSellingDealer(cboSELLING, 2)) & ","
        SQL = SQL & " EndUser = " & N2Str2Null(SetEndUser(cboEND, 2)) & ","
        SQL = SQL & " FIN_TYPE = " & vFINTYPE & ","
        SQL = SQL & " FIN_COMP = " & VFINNAME & ","
        SQL = SQL & " FIN_EXP_DATE = " & VFEXPDATE & ","
        SQL = SQL & " INS_TYPE = " & vINSTYPE & ","
        SQL = SQL & " INS_COMP = " & vINSNAME & ","
        SQL = SQL & " INS_EXP_DATE = " & VIEXPDATE
        SQL = SQL & " WHERE ID = " & labid.Caption

        gconDMIS.Execute (SQL)
        SQL_STATEMENT = SQL
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "CUSTOMER VEHICLE", SQL_STATEMENT, labid, "", "PLATE NO: " & Null2String(vPLATENO), "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        
        Call ShowSuccessFullyUpdated
    End If

    gconDMIS.CommitTrans
    Call cmdCancel_Click

    Dim TS                                             As String
    TS = txtSearch
    txtSearch.Text = ""
    Call InitializeRC

    txtSearch.Text = TS
    txtSearch.SetFocus
    
    Exit Sub
ERROR_MSG:
    MsgBox "Error Number : " & Err.Number & vbCrLf & _
        "Error Description : " & Err.Description, vbExclamation, "Error"
    Err.Clear
    gconDMIS.RollbackTrans
    Screen.MousePointer = 0
End Sub

Private Sub cmdx_Click()
    picDPLATE.Visible = False
    picDPLATE.ZOrder 1

    txtSearch.Enabled = True
    rptVEH.Enabled = True
    cmdPrint.Enabled = True
    cmdExit.Enabled = True

    VSEARCHMODE = "NO"
    txtSearch.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3:
            If rptVEH.Enabled = True Then
                txtSearch.SetFocus
            End If

        Case vbKeyEscape:
            If picEDIT.Visible = True Then
                Call cmdCancel_Click
            End If

        Case vbKeyF4:
            If Not VSEARCHMODE = "YES" Then
                picDPLATE.Visible = True
                picDPLATE.ZOrder 0
                Call DisplayDuplicatePlate
                VSEARCHMODE = "YES"

                txtSearch.Enabled = False
                rptVEH.Enabled = False
                cmdPrint.Enabled = False
                cmdExit.Enabled = False
            End If

        Case vbKeyF5:
            If Not VSEARCHMODE = "YES" Then
                txtSearch.Text = ""
                'Call displayVehicle
                Call InitializeRC
            End If

        Case vbKeyF6:
            If Not VSEARCHMODE = "YES" Then
                picDVIN.Visible = True
                picDVIN.ZOrder 0
                Call DisplayDuplicateVIN
                VSEARCHMODE = "YES"

                txtSearch.Enabled = False
                rptVEH.Enabled = False
                cmdPrint.Enabled = False
                cmdExit.Enabled = False
            End If
        
        Case vbKeyF7:
            If COMPANY_CODE = "HAI" Then
                If picEDIT.Visible = False Then
                    MessagePop InfoFriend, "Info", "Choose first a Vehicle"
                    Exit Sub
                End If
                If picEDIT.Enabled = False Then Exit Sub
                
                If Module_Access(LOGID, "INPUT LOYALTY NO", "SYSTEM") = False Then Exit Sub
                
                PICLoyaltyID.Visible = True
                PICLoyaltyID.ZOrder 0
                picEDIT.Enabled = False
                txtLoyalID.SetFocus
            Else
                MessagePop InfoFriend, "Module Info.", "This module is not supported by your dealer. For more information Kindly contact Netspeed Software Inc. about this Module"
                Exit Sub
            End If
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If AddorEdit = "ADD" And picEDIT.Visible = True Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CUSTOMER VEHICLE MASTER FILE)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "CUSTOMER VEHICLE", "")

    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 11
    Call FillCustomer
    Call FillColorS

    Call FillCboMoreYear(cboYear)
    cboYear.AddItem " "

    Call FillMake
    Call FillMOdelCode
    Call FillModel
    Call FillSellingDealer

    'Call displayVehicle
    Call InitializeRC
    Screen.MousePointer = 0
End Sub

Private Sub lsvDPLATE_DblClick()
    If lsvDPLATE.ListItems.Count = 0 Then Exit Sub

    Dim Index                                          As Integer
    Dim VPLATE                                         As String
    Index = lsvDPLATE.SelectedItem.Index

    VPLATE = lsvDPLATE.ListItems(Index).Text
    Call displayDuplicatePlateNo(VPLATE)

    Call cmdx_Click
End Sub

Private Sub lsvDVIN_DblClick()
    If lsvDVIN.ListItems.Count = 0 Then Exit Sub

    Dim Index                                          As Integer
    Dim vVin                                           As String
    Index = lsvDVIN.SelectedItem.Index

    vVin = lsvDVIN.ListItems(Index).Text
    Call displayDuplicateVinNo(vVin)

    Call cmd1_Click
End Sub

Private Sub rptVEH_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem)
    'On Error Resume Next
    Dim Index                                          As Long
    Dim vID                                            As Long
    Dim vRONO                                          As String

    If Row.Record Is Nothing Then: Exit Sub
    If Function_Access(LOGID, "Acess_EDIT", "CUSTOMER VEHICLE") = False Then Exit Sub
    AddorEdit = "EDIT"
    cmdDelete.Visible = True
    labid.Caption = Null2String(Row.Record(9).Value)
    cboNAME.Text = FindEndUser(Null2String(Row.Record(0).Value))
    txtPLATENO.Text = Null2String(Row.Record(2).Value)
    LBLOLDPLATE.Caption = Null2String(Row.Record(2).Value)

    txtCONDNO.Text = Null2String(Row.Record(3).Value)
    txtVIN.Text = Null2String(Row.Record(4).Value)
    LBLOLDVIN.Caption = Null2String(Row.Record(4).Value)

    txtDesc.Text = Null2String(Row.Record(5).Value)
    txtENGINE.Text = Null2String(Row.Record(6).Value)

    'txtDDATE.Text = Null2Date(Row.Record(7).Value)
    txtDDATE.Value = Null2Date(Row.Record(7).Value)
    'txtDDATE.Text = Null2Date(Row.Record(8).Value)
    txtDDATE.Value = Null2Date(Row.Record(8).Value)

    Call FillOtherInfo(labid)
    picEDIT.Visible = True
    picEDIT.ZOrder 0

    txtSearch.Enabled = False
    rptVEH.Enabled = False
    cmdPrint.Enabled = False
    cmdExit.Enabled = False
End Sub

Private Sub Timer1_Timer()
    If lblloyal.ForeColor = vbRed Then
        lblloyal.ForeColor = vbBlue
    Else
        lblloyal.ForeColor = vbRed
    End If
End Sub

Private Sub txtDDATE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890/", KeyAscii)
    End If
End Sub

Private Sub txtFDATE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890/", KeyAscii)
    End If
End Sub

Private Sub txtFDATE_LostFocus()
    '    If Not txtFDATE.Text = "" Then
    '        If IsDate(txtFDATE) = False Then
    '            ShowIsRequiredMsg "Invalid Date Format"
    '            txtFDATE.SetFocus
    '            Exit Sub
    '        End If
    '    End If
End Sub

Private Sub txtIDATE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890/", KeyAscii)
    End If
End Sub

Private Sub txtIDATE_LostFocus()
    '    If Not txtIDATE = "" Then
    '        If IsDate(txtIDATE) = False Then
    '            ShowIsRequiredMsg "Invalid Date Format"
    '            txtIDATE.Text = ""
    '            Exit Sub
    '        End If
    '    End If
End Sub

Private Sub txtSDATE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890/", KeyAscii)
    End If
End Sub

Private Sub txtSEARCH_Change()
    rptVEH.FilterText = txtSearch.Text
    rptVEH.Populate
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.BackColor = &HC0FFFF
End Sub

Private Sub txtSearch_LostFocus()
    txtSearch.BackColor = vbWhite
End Sub

Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False

    Dim ar()                                           As String
    Dim cWidth                                         As Long
    Dim i                                              As Integer
    Dim scwidth                                        As Long
    ar = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For i = LBound(ar) To UBound(ar)
            If i <= grd.ColumnHeaders.Count Then
                scwidth = cWidth * (CDec(ar(i)) / 100)
                grd.ColumnHeaders(i + 1).Width = scwidth
            End If
        Next
    ElseIf TypeOf grd Is ReportControl Then
        For i = LBound(ar) To UBound(ar)
            If i < grd.Columns.Count Then
                scwidth = cWidth * (CDec(ar(i)) / 100)
                grd.Columns(i).Width = scwidth
            End If
        Next

    End If

    Erase ar
    grd.Visible = True
End Sub

Sub InitializeRC()
    With rptVEH
        'DATE SOLD, DATE DELIVERY, , MAKE")
        .Columns.DeleteAll
        .Columns.Add 0, "Code", 0, True::         .Columns(0).Alignment = xtpAlignmentLeft:       .Columns(0).AllowRemove = False
        .Columns.Add 1, "Name", 250, True::         .Columns(1).Alignment = xtpAlignmentLeft:       .Columns(1).AllowRemove = False
        .Columns.Add 2, "Plate no", 60, True:      .Columns(2).Alignment = xtpAlignmentCenter:       .Columns(2).AllowRemove = False
        .Columns.Add 3, "Cond. no", 60, True:       .Columns(3).Alignment = xtpAlignmentCenter:       .Columns(3).AllowRemove = False
        .Columns.Add 4, "Vin no", 130, True:         .Columns(4).Alignment = xtpAlignmentLeft:       .Columns(4).AllowRemove = False
        .Columns.Add 5, "Description", 120, True:   .Columns(5).Alignment = xtpAlignmentLeft:       .Columns(5).AllowRemove = False
        .Columns.Add 6, "Engine no", 90, True:     .Columns(6).Alignment = xtpAlignmentCenter:       .Columns(6).AllowRemove = False
        .Columns.Add 7, "Date sold", 80, True:      .Columns(7).Alignment = xtpAlignmentCenter:       .Columns(7).AllowRemove = False
        .Columns.Add 8, "Date Delivery", 90, True:   .Columns(8).Alignment = xtpAlignmentCenter:       .Columns(8).AllowRemove = False
        .Columns.Add 9, "", 0, True:                .Columns(9).Alignment = xtpAlignmentLeft:       .Columns(9).AllowRemove = False
        .Columns.Add 10, "Make", 80, True:            .Columns(10).Alignment = xtpAlignmentLeft:       .Columns(10).AllowRemove = False
        
        .GroupsOrder.Add .Columns(0)
        .Columns(0).Visible = False
        .Columns(9).Visible = False
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
    End With
    
    
    Dim RSUPLOAD                                        As New ADODB.Recordset
    Dim REC                                             As XtremeReportControl.ReportRecord
    Dim XXX                                             As String
    
    XXX = "SELECT  CUSCDE, NIYM, PLATE_NO, VCOND_NO, VIN, DESCRIPTION, ENGINE, D_SOLD, DEL_DATE, ID, MAKE FROM CSMS_CUSVEH"
    Set RSUPLOAD = gconDMIS.Execute(XXX)
    rptVEH.Records.DeleteAll
    While Not RSUPLOAD.EOF
        Set REC = rptVEH.Records.Add
        REC.AddItem (Trim(RSUPLOAD!CUSCDE))
        REC.AddItem (Trim(RSUPLOAD!NIYM))
        REC.AddItem (Trim(RSUPLOAD!PLATE_NO))
        REC.AddItem (Trim(RSUPLOAD!VCOND_NO))
        REC.AddItem (Trim(RSUPLOAD!Vin))
        REC.AddItem (Trim(RSUPLOAD!Description))
        REC.AddItem (Trim(RSUPLOAD!Engine))
        REC.AddItem (Trim(RSUPLOAD!D_SOLD))
        REC.AddItem (Trim(RSUPLOAD!DEL_DATE))
        REC.AddItem (Trim(RSUPLOAD!ID))
        REC.AddItem (Trim(RSUPLOAD!Make))
        
        RSUPLOAD.MoveNext
        Set REC = Nothing
    Wend
    rptVEH.Populate
End Sub

