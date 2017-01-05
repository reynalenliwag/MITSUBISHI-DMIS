VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Log_TestDriveAppointment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ":::Test Drive"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogTestDriveAppointment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5325
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   1785
      Width           =   2415
      Begin VB.OptionButton optDate 
         Caption         =   "Test Vehicles Models"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   405
         Width           =   2265
      End
      Begin VB.OptionButton optAcctName 
         Caption         =   "Search By Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   135
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   90
         TabIndex        =   3
         Top             =   720
         Width           =   2310
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4125
         Left            =   45
         TabIndex        =   4
         Top             =   1170
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   7276
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1755
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   8055
      TabIndex        =   26
      Top             =   0
      Width           =   8055
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   2490
         Top             =   1110
      End
      Begin VB.TextBox txtEntityName 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   32
         Top             =   210
         Width           =   4935
      End
      Begin VB.TextBox txtEntityContactperson 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   31
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox txtEntityAddress 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   1230
         Width           =   4665
      End
      Begin VB.TextBox txtEntityPhone 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5070
         TabIndex        =   29
         Top             =   210
         Width           =   2670
      End
      Begin VB.TextBox txtEntityMobile 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5070
         TabIndex        =   28
         Top             =   720
         Width           =   2370
      End
      Begin VB.TextBox txtEntityEmail 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5100
         TabIndex        =   27
         Top             =   1260
         Width           =   2790
      End
      Begin VB.Label labEntityName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CUSTOMER NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   38
         Top             =   0
         Width           =   1410
      End
      Begin VB.Label labEntityAddress 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   37
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CONTACT PERSON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   36
         Top             =   510
         Width           =   1470
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "PHONE NUMBER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4830
         TabIndex        =   35
         Top             =   0
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "EMAIL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4830
         TabIndex        =   34
         Top             =   1020
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "MOBILE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4830
         TabIndex        =   33
         Top             =   510
         Width           =   1230
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   240
         X2              =   7980
         Y1              =   1710
         Y2              =   1710
      End
   End
   Begin VB.PictureBox picDataEntry 
      BorderStyle     =   0  'None
      Height          =   4545
      Left            =   2415
      ScaleHeight     =   4545
      ScaleWidth      =   6135
      TabIndex        =   5
      Top             =   1770
      Width           =   6135
      Begin VB.TextBox txtStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   150
         Width           =   4110
      End
      Begin VB.ComboBox cboColor 
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
         Left            =   1410
         TabIndex        =   40
         Text            =   "cboClassification"
         Top             =   1830
         Width           =   4125
      End
      Begin VB.TextBox txtInterest 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   1395
         TabIndex        =   19
         Top             =   2610
         Width           =   4125
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
         Left            =   1425
         TabIndex        =   9
         Text            =   "cboAttendingSE"
         Top             =   1020
         Width           =   4125
      End
      Begin VB.TextBox txtFeedBack 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   1410
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   3540
         Width           =   4125
      End
      Begin VB.ComboBox cboClassification 
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
         Left            =   1425
         TabIndex        =   12
         Text            =   "cboClassification"
         Top             =   1470
         Width           =   4125
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
         Left            =   1425
         TabIndex        =   8
         Text            =   "cboVehicles"
         Top             =   615
         Width           =   4125
      End
      Begin MSComCtl2.DTPicker txtStartTime 
         Height          =   360
         Left            =   3090
         TabIndex        =   15
         Top             =   2220
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
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
         CustomFormat    =   "hh:mm tt"
         Format          =   20643842
         CurrentDate     =   39084
      End
      Begin MSComCtl2.DTPicker txtdtDate 
         Height          =   360
         Left            =   1395
         TabIndex        =   14
         Top             =   2220
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   635
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
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   20643841
         CurrentDate     =   39084
      End
      Begin MSComCtl2.DTPicker txtEndTime 
         Height          =   360
         Left            =   4365
         TabIndex        =   16
         Top             =   2220
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   635
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
         CustomFormat    =   "hh:mm tt"
         Format          =   20643842
         CurrentDate     =   39084
      End
      Begin VB.TextBox txtCode 
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
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3030
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   720
         TabIndex        =   41
         Top             =   150
         Width           =   555
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   870
         TabIndex        =   39
         Top             =   1890
         Width           =   450
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   450
         TabIndex        =   13
         Top             =   2265
         Width           =   855
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SAE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   8
         Left            =   960
         TabIndex        =   11
         Top             =   1065
         Width           =   345
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   585
         Index           =   6
         Left            =   135
         TabIndex        =   18
         Top             =   2580
         Width           =   1170
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Feed Back"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   420
         TabIndex        =   17
         Top             =   3510
         Width           =   885
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Evaulation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   435
         TabIndex        =   10
         Top             =   1515
         Width           =   870
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Model Descript"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   6
         Top             =   660
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   8055
      TabIndex        =   21
      Top             =   6300
      Width           =   8055
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   2670
         ScaleHeight     =   900
         ScaleWidth      =   5490
         TabIndex        =   43
         Top             =   0
         Width           =   5490
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   4560
            MouseIcon       =   "LogTestDriveAppointment.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "LogTestDriveAppointment.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Exit Window"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3870
            MouseIcon       =   "LogTestDriveAppointment.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "LogTestDriveAppointment.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Delete Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   3180
            MouseIcon       =   "LogTestDriveAppointment.frx":11FF
            MousePointer    =   99  'Custom
            Picture         =   "LogTestDriveAppointment.frx":1351
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Edit Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   2490
            MouseIcon       =   "LogTestDriveAppointment.frx":16AD
            MousePointer    =   99  'Custom
            Picture         =   "LogTestDriveAppointment.frx":17FF
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Add Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   1860
            MouseIcon       =   "LogTestDriveAppointment.frx":1B12
            MousePointer    =   99  'Custom
            Picture         =   "LogTestDriveAppointment.frx":1C64
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Find a Record"
            Top             =   60
            Width           =   645
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1170
            MouseIcon       =   "LogTestDriveAppointment.frx":1F5E
            MousePointer    =   99  'Custom
            Picture         =   "LogTestDriveAppointment.frx":20B0
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Move to Next Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   480
            MouseIcon       =   "LogTestDriveAppointment.frx":2408
            MousePointer    =   99  'Custom
            Picture         =   "LogTestDriveAppointment.frx":255A
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Move to Previous Record"
            Top             =   60
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   6480
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   23
         Top             =   -15
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   750
            MouseIcon       =   "LogTestDriveAppointment.frx":28B9
            MousePointer    =   99  'Custom
            Picture         =   "LogTestDriveAppointment.frx":2A0B
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Cancel"
            Top             =   65
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   60
            MouseIcon       =   "LogTestDriveAppointment.frx":2D49
            MousePointer    =   99  'Custom
            Picture         =   "LogTestDriveAppointment.frx":2E9B
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Save this Record"
            Top             =   65
            Width           =   705
         End
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         Height          =   510
         Left            =   270
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSMIS_Log_TestDriveAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PROSPECTID                                                        As Long
Dim ScheduleID                                                        As Long
Dim RS                                                                As ADODB.Recordset

Sub FillSearchGrid(XXX As String)
    Dim TEMPRS                                                        As ADODB.Recordset

    ListView1.Enabled = False

    If optAcctName.Value = True Then
        Set TEMPRS = gconDMIS.Execute("SELECT  convert(varchar, StartDateTime,101), VehicleModel , ScheduleID FROM CRIS_TestDriveSchedules where  ProspectID=" & PROSPECTID & " AND  Convert(varchar, StartDateTime , 101)  like  '" & ReplaceQuote(XXX) & "%' order by 1  asc")
    Else
        Set TEMPRS = gconDMIS.Execute("SELECT  convert(varchar, StartDateTime,101), VehicleModel , ScheduleID FROM CRIS_TestDriveSchedules where ProspectID=" & PROSPECTID & " AND   VehicleModel  like  '" & ReplaceQuote(XXX) & "%' order by 1  asc")
    End If

    If Not TEMPRS.EOF And Not TEMPRS.BOF Then
        ListView1.Enabled = True
    End If

    flex_FillListView TEMPRS, ListView1



End Sub

Sub SetEntityDetails(xProspectID As Long, xCUSCODE As String)
    Dim TEMPRS                                                        As ADODB.Recordset
    txtEntityAddress = ""
    txtEntityContactperson = ""
    txtEntityEmail = ""
    txtEntityMobile = ""
    txtEntityName = ""
    txtEntityPhone = ""

    If xProspectID = 0 Then
        labEntityName = "CUSTOMER NAME"
        Set TEMPRS = gconDMIS.Execute("Select CUSTOMERNAME as [Name], CONTACTPERSON, PHONE, MOBILE, ADDRESS, EMAIL from CRIS_VW_ALLPROFILE WHERE CUSCDE=" & N2Str2Null(xCUSCODE))
    Else
        labEntityName = "PROSPECT NAME"
        Set TEMPRS = gconDMIS.Execute("Select ACCTNAME As [NAME], CONTACTPERSON, TELEPHONE as PHONE , MOBILE, ADDRESS , EMAIL  from CRIS_PROSPECTS WHERE PROSPECTID=" & N2Str2Null(xProspectID))
    End If

    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        txtEntityAddress = Null2String(TEMPRS!Address)
        txtEntityContactperson = Null2String(TEMPRS!ContactPerson)
        txtEntityEmail = Null2String(TEMPRS!EMAIL)
        txtEntityMobile = Null2String(TEMPRS!Mobile)
        txtEntityName = Null2String(TEMPRS!Name)
        txtEntityPhone = Null2String(TEMPRS!Phone)
    End If
    Set TEMPRS = Nothing
End Sub

Sub InitMemVars()
    TXTCODE = ""
    txtdtDate = LOGDATE
    txtStartTime = FormatDateTime(LOGDATE, vbLongTime)
    txtEndTime = FormatDateTime(DateAdd("h", 1, LOGTIME), vbLongTime)
    txtFeedBack = ""
    txtSEARCH = ""
    txtInterest = ""
    cboAttendingSE.ListIndex = -1
End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    'SELECT ScheduleID, ProspectID, VehicleModel, VehicleCode, SAE, PossibleNextVisit, NextVisitNotes, Interests, FeedBack, StartDateTime, EndDateTime, Classification, ClosedDate FROM DMIS.dbo.CRIS_TestDriveSchedules
    RS.Open "SELECT * FROM CRIS_TestDriveSchedules Where ProspectID=" & PROSPECTID & "Order BY ScheduleID desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub SetDefaultChoice()
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("Select * from CRIS_PROSPECTS WHERE PROSPECTID=" & PROSPECTID)
    If Not (TEMPRS.BOF Or TEMPRS.EOF) Then
        cboAttendingSE.ListIndex = SelectCombo(cboAttendingSE, Null2String(TEMPRS!SAE))
        cboVehicles.ListIndex = SelectCombo(cboVehicles, Null2String(TEMPRS!Variant))

    End If
End Sub

Sub StoreMemVars()
    'ScheduleID, ProspectID, VehicleModel, VehicleCode, SAE, PossibleNextVisit, NextVisitNotes, Interests, FeedBack,
    'StartDateTime, EndDateTime, Classification, ClosedDate
    If Not RS.EOF And Not RS.BOF Then
        ScheduleID = RS!ScheduleID
        TXTCODE = Null2String(RS!vehiclecode)
        cboVehicles.ListIndex = SelectCombo(cboVehicles, Null2String(RS!vehiclemodel))
        txtdtDate = DateValue(RS!StartDateTime)

        txtStartTime = TimeValue(RS!StartDateTime)
        txtEndTime = TimeValue(RS!EndDateTime)

        txtFeedBack = Null2String(RS!FeedBack)
        txtInterest = Null2String(RS!Interests)
        txtStatus = Null2String(RS!STATUS)
        cboAttendingSE.ListIndex = SelectCombo(cboAttendingSE, Null2String(RS!SAE))
        If LOGSAE <> "" Then
            cboAttendingSE.Enabled = False
        Else
            cboAttendingSE.Enabled = True
        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub UpdateLog()
    Dim TSQL                                                          As String
    TSQL = " DECLARE @DT DATETIME" & vbCrLf
    TSQL = TSQL & " SELECT   @DT=MAX(CAST(CONVERT(VARCHAR, STARTDATETIME ,101) + ' ' + CONVERT(VARCHAR, ENDDATETIME  ,114)  AS SMALLDATETIME)) FROM CRIS_TESTDRIVESCHEDULES  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
    TSQL = TSQL & " BEGIN " & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGTESTDRIVE=@DT , HITCOUNTER=1  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End " & vbCrLf
    TSQL = TSQL & " Else " & vbCrLf
    TSQL = TSQL & " BEGIN" & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGTESTDRIVE=NULL  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End"
    gconDMIS.Execute (TSQL)
End Sub

Sub FillColori()
    Dim rsColor                                                       As New ADODB.Recordset
    Dim SQL                                                           As String

    SQL = "SELECT Color_desc FROM All_Color"

    Set rsColor = New ADODB.Recordset
    Set rsColor = gconDMIS.Execute(SQL)

    cboColor.Clear

    Do While Not rsColor.EOF
        cboColor.AddItem Null2String(rsColor!color_desc)
        rsColor.MoveNext
    Loop
    Set rsColor = Nothing
End Sub

''''''CALLS
Friend Sub AddTestDriveAppointment(xProspectID As Long)
    PROSPECTID = xProspectID
    ScheduleID = 0
End Sub

Friend Sub EditTestDriveAppointment(xProspectID As Long, xScheduleID As Long)
    PROSPECTID = xProspectID
    ScheduleID = xScheduleID
End Sub

Private Sub cboVehicles_Click()
    If cboVehicles.ListIndex = -1 Then Exit Sub
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT CODE FROM  SMIS_MRRINV  WHERE DESCRIPT='" & LTrim(RTrim(cboVehicles.Text)) & "'")
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        TXTCODE = Null2String(TEMPRS!CODE)
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "LOG TEST DRIVE") = False Then Exit Sub
    On Error GoTo ErrorCode:

    ScheduleID = 0
    InitMemVars
    SetDefaultChoice
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    PICSEARCH.Enabled = False
    On Error Resume Next
    'cboVehicles.SetFocus
    txtStatus.Text = "For Approval"





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False
    PICSEARCH.Enabled = True
    ScheduleID = 0
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()

    If Function_Access(LOGID, "Acess_DELETE", "LOG TEST DRIVE") = False Then Exit Sub

    On Error GoTo ErrorCode:

    If IsDate(RS!ClosedDate) = True Then
        MessagePop RecLocekd, "Record In Use", " Current Schedule Information is in Use  or Been Closed.... Cannot Delete The Record"
        Exit Sub
    End If
    If ShowConfirmDelete = True Then
        SQL_STATEMENT = "delete from CRIS_TestDriveSchedules where ScheduleID = " & ScheduleID

        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "X", "TEST DRIVE VEHICLES", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""

        ShowDeletedMsg
        UpdateLog
        FillSearchGrid txtSEARCH
        rsRefresh
        StoreMemVars
        If FormExist("MainForm") Then
            MainForm.ShowStatus PROSPECTID
        End If

    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "LOG TEST DRIVE") = False Then Exit Sub
    On Error GoTo ErrorCode:

    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    PICSEARCH.Enabled = False
    On Error Resume Next
    cboVehicles.SetFocus





    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub cmdNext_Click()
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdPrevious_Click()
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdSave_Click()
    Dim SQL                                                           As String
    Dim t1 As String, t2                                              As String


    On Error GoTo ErrorCode:

    If Trim(cboClassification) = "" Then
        ShowIsRequiredMsg "Evalutaion"
        On Error Resume Next
        cboClassification.SetFocus
        Exit Sub
    End If



    If ScheduleID <= 0 Then
        SQL = "INSERT INTO CRIS_TestDriveSchedules " & _
            " (ProspectID, VehicleModel, VehicleCode, SAE, PossibleNextVisit, NextVisitNotes, Interests, FeedBack, StartDateTime,EndDateTime, Classification,color,status) " & _
            " VALUES(@ProspectID, @VehicleModel, @VehicleCode , @SAE, @PossibleNextVisit, @NextVisitNotes, @Interests, @FeedBack, @StartDateTime, @EndDateTime, @Classification,'" & cboColor.Text & "','" & txtStatus.Text & "') " & vbCrLf & " SELECT @@IDENTITY"

        SQL_STATEMENT = SQL
        NEW_LogAudit "A", "TEST DRIVE VEHICLES", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""

        gconDMIS.Execute "update cris_mrrinv set HITCOUNTER=(ISNULL(HITCOUNTER,0) + 1) where CODE=" & N2Str2Null(TXTCODE)


    Else
        SQL = "Update CRIS_TestDriveSchedules " & _
            " SET  ProspectID=@ProspectID, VehicleModel= @VehicleModel, VehicleCode= @VehicleCode , SAE=@SAE, PossibleNextVisit=@PossibleNextVisit, NextVisitNotes=@NextVisitNotes, Interests=@Interests, StartDateTime=@StartDateTime, EndDateTime= @EndDateTime, Classification=@Classification,Color='" & cboColor.Text & "'" & _
            "  WHERE ScheduleID=@ScheduleID"

        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "TEST DRIVE VEHICLES", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""

    End If




    t1 = DateValue(txtdtDate) & " " & TimeValue(txtStartTime)
    t2 = DateValue(txtdtDate) & " " & TimeValue(txtEndTime)


    SQL = Replace(SQL, "@ScheduleID", ScheduleID)
    SQL = Replace(SQL, "@ProspectID", PROSPECTID)
    SQL = Replace(SQL, "@VehicleModel", N2Str2Null(cboVehicles.Text))
    SQL = Replace(SQL, "@VehicleCode", N2Str2Null(TXTCODE.Text))
    SQL = Replace(SQL, "@SAE", N2Str2Null(cboAttendingSE.Text))
    SQL = Replace(SQL, "@PossibleNextVisit", N2Str2Null(""))
    SQL = Replace(SQL, "@NextVisitNotes", N2Str2Null(""))
    SQL = Replace(SQL, "@Interests", N2Str2Null(txtInterest))
    SQL = Replace(SQL, "@FeedBack", N2Str2Null(txtFeedBack))
    SQL = Replace(SQL, "@StartDateTime", N2Str2Null(t1))
    SQL = Replace(SQL, "@EndDateTime", N2Str2Null(t2))
    SQL = Replace(SQL, "@Classification", N2Str2Null(cboClassification))




    Dim TEMPRS                                                        As ADODB.Recordset

    Set TEMPRS = gconDMIS.Execute(SQL)
    gconDMIS.Execute ("update CRIS_PROSPECTs SET LogTestDrive=" & N2Str2Null(t1) & " where prospectid=" & PROSPECTID)

    If ScheduleID <= 0 Then
        MessagePop RecSaveOk, "Record Added ", "New Test Drive Schedule Sucessfully Added", 500, 1
    Else
        MessagePop RecSaveOk, "RecordSaved", "Test Drive Schedule Sucessfully Updated", 500, 1
    End If

    Set TEMPRS = TEMPRS.NextRecordset
    If Not TEMPRS Is Nothing Then
        ScheduleID = TEMPRS.Collect(0)
    End If

    UpdateLog
    rsRefresh
    RS.Find ("ScheduleID=" & ScheduleID)
    FillSearchGrid txtSEARCH
    Set TEMPRS = Nothing
    cmdCancel.Value = True

    If FormExist("MainForm") Then
        MainForm.ShowStatus PROSPECTID
    End If
    Exit Sub
ErrorCode:
    ShowVBError

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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (TEST DRIVE VEHICLES)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(PROSPECTID), "TEST DRIVE VEHICLES")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitMemVars
    InitVars
    rsRefresh
    StoreMemVars
    SetDefaultChoice
    SetEntityDetails PROSPECTID, vbNullString
    FillColori
End Sub

Private Sub InitVars()
    Combo_Loadval cboVehicles, gconDMIS.Execute("Select DISTINCT Descript  from ALL_MODEL")
    Call FillCombo("SELECT ID, [Name] from SMIS_vw_Srep ORDER BY 2", 0, 1, cboAttendingSE)
    Combo_Loadval cboClassification, gconDMIS.Execute("Select datadesc from CRIS_vw_masterPullDown where MasterDesc ='Customer Classification'")


    Call AddColumnHeader("Date, Model ", ListView1)
    Call ResizeColumnHeader(ListView1, "35,60")
    FillSearchGrid txtSEARCH
    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False
    PICSEARCH.Enabled = True
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListView1
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub LISTVIEW1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RS.MoveFirst
    RS.Find ("ScheduleID=" & Item.ListSubItems(2).Text)
    StoreMemVars
End Sub

Private Sub optAcctName_Click()
    FillSearchGrid txtSEARCH
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub optDate_Click()
    FillSearchGrid txtSEARCH
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub Timer1_Timer()
    If txtStatus.ForeColor = &HC0& Then
        txtStatus.ForeColor = &HC0C0&
    Else
        txtStatus.ForeColor = &HC0&
    End If
End Sub

Private Sub txtSEARCH_Change()
    FillSearchGrid txtSEARCH

End Sub

