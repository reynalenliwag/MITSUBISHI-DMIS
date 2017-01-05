VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Log_SalesAppointment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Sales Appointment"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SalesAppointment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   7500
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1755
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   7500
      TabIndex        =   40
      Top             =   0
      Width           =   7500
      Begin VB.TextBox txtEntityName 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   60
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   210
         Width           =   4935
      End
      Begin VB.TextBox txtEntityContactperson 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   60
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox txtEntityAddress 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Text            =   "SalesAppointment.frx":08CA
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox txtEntityPhone 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   5070
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   210
         Width           =   2370
      End
      Begin VB.TextBox txtEntityMobile 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   5070
         TabIndex        =   42
         Text            =   "09175041620"
         Top             =   720
         Width           =   2370
      End
      Begin VB.TextBox txtEntityEmail 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   5070
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   1260
         Width           =   2370
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         Left            =   5070
         TabIndex        =   49
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
         Left            =   5070
         TabIndex        =   48
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
         Left            =   5070
         TabIndex        =   47
         Top             =   510
         Width           =   1230
      End
   End
   Begin VB.PictureBox picDataEntry 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4320
      Left            =   2535
      ScaleHeight     =   4320
      ScaleWidth      =   5280
      TabIndex        =   5
      Top             =   1755
      Width           =   5280
      Begin VB.TextBox txtNotes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1380
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   3270
         Width           =   3570
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
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2040
         Width           =   1965
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
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1650
         Width           =   3570
      End
      Begin VB.TextBox txtModel 
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
         Height          =   345
         Left            =   1350
         TabIndex        =   19
         Top             =   2040
         Width           =   1515
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
         Left            =   1350
         TabIndex        =   22
         Top             =   2430
         Width           =   3570
      End
      Begin VB.ComboBox cboTerms 
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
         Left            =   1350
         TabIndex        =   23
         Top             =   2820
         Width           =   1740
      End
      Begin VB.ComboBox cboImportance 
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
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   615
         Width           =   3570
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
         Left            =   1350
         TabIndex        =   6
         Text            =   "cboAttendingSE"
         Top             =   180
         Width           =   3570
      End
      Begin MSComCtl2.DTPicker txtExpectedPurchase 
         Height          =   360
         Left            =   3150
         TabIndex        =   25
         Top             =   2820
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   54591491
         CurrentDate     =   39171
      End
      Begin MSComCtl2.DTPicker txtStartTime 
         Height          =   360
         Left            =   2640
         TabIndex        =   14
         Top             =   1230
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm tt"
         Format          =   54591491
         UpDown          =   -1  'True
         CurrentDate     =   39084
      End
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   360
         Left            =   1350
         TabIndex        =   11
         Top             =   1230
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   54591489
         CurrentDate     =   39171
      End
      Begin MSComCtl2.DTPicker txtEndTime 
         Height          =   360
         Left            =   3780
         TabIndex        =   15
         Top             =   1230
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm tt"
         Format          =   54591491
         UpDown          =   -1  'True
         CurrentDate     =   39084
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   225
         Index           =   6
         Left            =   825
         TabIndex        =   53
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Code/Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   330
         TabIndex        =   18
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   1680
         Width           =   1320
      End
      Begin VB.Label lblCap 
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
         Left            =   465
         TabIndex        =   9
         Top             =   1290
         Width           =   855
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "TIME FROM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   2640
         TabIndex        =   12
         Top             =   1020
         Width           =   945
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "TIME TO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   3780
         TabIndex        =   13
         Top             =   1020
         Width           =   690
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
         Height          =   225
         Index           =   4
         Left            =   870
         TabIndex        =   21
         Top             =   2490
         Width           =   450
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         Caption         =   "Expected Terms / Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   75
         TabIndex        =   24
         Top             =   2790
         Width           =   1245
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Importance"
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
         Left            =   345
         TabIndex        =   7
         Top             =   645
         Width           =   975
      End
      Begin VB.Label lblCap 
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
         Left            =   975
         TabIndex        =   10
         Top             =   240
         Width           =   345
      End
   End
   Begin VB.PictureBox picSearch 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4320
      Left            =   0
      ScaleHeight     =   4320
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   1755
      Width           =   2535
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   60
         TabIndex        =   3
         Top             =   720
         Width           =   2460
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
         Left            =   150
         TabIndex        =   1
         Top             =   135
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Vehicles Models"
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
         Left            =   150
         TabIndex        =   2
         Top             =   405
         Width           =   2265
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3195
         Left            =   45
         TabIndex        =   4
         Top             =   1110
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   5636
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
   Begin VB.PictureBox Picture5 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   0
      ScaleHeight     =   1020
      ScaleWidth      =   7500
      TabIndex        =   27
      Top             =   6075
      Width           =   7500
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   2055
         ScaleHeight     =   900
         ScaleWidth      =   5490
         TabIndex        =   29
         Top             =   45
         Width           =   5490
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   4560
            MouseIcon       =   "SalesAppointment.frx":08D0
            MousePointer    =   99  'Custom
            Picture         =   "SalesAppointment.frx":0A22
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Exit Window"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3870
            MouseIcon       =   "SalesAppointment.frx":0D88
            MousePointer    =   99  'Custom
            Picture         =   "SalesAppointment.frx":0EDA
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Delete Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   3180
            MouseIcon       =   "SalesAppointment.frx":1205
            MousePointer    =   99  'Custom
            Picture         =   "SalesAppointment.frx":1357
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Edit Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   2490
            MouseIcon       =   "SalesAppointment.frx":16B3
            MousePointer    =   99  'Custom
            Picture         =   "SalesAppointment.frx":1805
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Add Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   1800
            MouseIcon       =   "SalesAppointment.frx":1B18
            MousePointer    =   99  'Custom
            Picture         =   "SalesAppointment.frx":1C6A
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Find a Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1110
            MouseIcon       =   "SalesAppointment.frx":1F64
            MousePointer    =   99  'Custom
            Picture         =   "SalesAppointment.frx":20B6
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Move to Next Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   420
            MouseIcon       =   "SalesAppointment.frx":240E
            MousePointer    =   99  'Custom
            Picture         =   "SalesAppointment.frx":2560
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Move to Previous Record"
            Top             =   60
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   5805
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   37
         Top             =   45
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   780
            MouseIcon       =   "SalesAppointment.frx":28BF
            MousePointer    =   99  'Custom
            Picture         =   "SalesAppointment.frx":2A11
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Cancel"
            Top             =   65
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   90
            MouseIcon       =   "SalesAppointment.frx":2D4F
            MousePointer    =   99  'Custom
            Picture         =   "SalesAppointment.frx":2EA1
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Save this Record"
            Top             =   65
            Width           =   705
         End
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         Height          =   510
         Left            =   270
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSMIS_Log_SalesAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PROSPECTID                                                        As Long
Dim AppointmentID                                                     As Long
Dim RS                                                                As ADODB.Recordset

Sub FillSearchGrid(XXX As String)
    Dim TEMPRS                                                        As ADODB.Recordset

    ListView1.Enabled = False

    If optAcctName.Value = True Then
        Set TEMPRS = gconDMIS.Execute("SELECT CONVERT( VARCHAR, STARTDATETIME,101), MODELDESCRIPT, APPOINTMENTID FROM CRIS_SALESAPPOINTMENTS WHERE  ProspectID=" & PROSPECTID & " AND  CONVERT(VARCHAR, STARTDATETIME, 101)  LIKE  '" & ReplaceQuote(XXX) & "%' ORDER BY 1  ASC")

    Else
        Set TEMPRS = gconDMIS.Execute("SELECT CONVERT( VARCHAR, STARTDATETIME,101), MODELDESCRIPT, APPOINTMENTID  FROM CRIS_SALESAPPOINTMENTS WHERE  ProspectID=" & PROSPECTID & " AND  MODELDESCRIPT LIKE  '" & ReplaceQuote(XXX) & "%' ORDER BY 2  ASC")
    End If

    If Not TEMPRS.EOF And Not TEMPRS Then
        ListView1.Enabled = True
        Listview_Loadval ListView1.ListItems, TEMPRS

    End If





End Sub

Sub InitData()

    Call FillCombo("SELECT DISTINCT Name from SMIS_vw_Srep  ORDER BY [name]", -1, 0, cboAttendingSE)
    Call FillCombo("Select DISTINCT 1, COLOR_DESC FROM ALL_COLOR ORDER BY COLOR_DESC", 0, 1, cboColors)
    Call FillCombo("select ID, DESCRIPT from ALL_MODEL", 0, 1, cboVehicles)
    With cboImportance
        .AddItem "Normal"
        .AddItem "High"
        .AddItem "Very High"
        .AddItem "Low"
        .ListIndex = 0
    End With
    With cboTerms
        .AddItem "Cash"
        .AddItem "Financing"
        .AddItem "Others"
        .ListIndex = 0
    End With
    AddColumnHeader "Date , EmailAddress", ListView1
    ResizeColumnHeader ListView1, "40,55"
    FillSearchGrid ""


End Sub

Sub InitMemVars()
    AppointmentID = 0
    cboAttendingSE = SAENAME
    txtStartTime = TimeValue("8:00AM")

    txtStartTime.MinDate = TimeValue(LOGTIME)
    txtStartTime.MaxDate = TimeValue(LOGTIME)
    txtDate = DateValue(LOGDATE)
    txtEndTime = TimeValue(LOGTIME)
    txtModel = ""
    txtModelCode = ""
    cboVehicles.ListIndex = -1
    cboColors.ListIndex = -1
    cboTerms.ListIndex = -1
    txtExpectedPurchase = DateValue(LOGDATE)
    txtNotes = ""

End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM CRIS_SalesAppointments Where ProspectID=" & PROSPECTID & " Order BY StartDateTime desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not RS.EOF And Not RS.BOF Then
        'AppointmentID, ProspectID, SAE, StartDateTime, EndDateTime, ClosedDate, Model, ModelDescript, ModelCode,Color, Terms, ExpectedPurchase, Notes
        AppointmentID = RS!AppointmentID
        PROSPECTID = RS!PROSPECTID
        cboAttendingSE = Null2String(RS!SAE)
        txtStartTime = TimeValue(RS!StartDateTime)
        txtEndTime = TimeValue(RS!EndDateTime)
        txtModel = Null2String(RS!Model)
        txtDate = Null2String(RS!StartDateTime)
        txtModelCode = Null2String(RS!ModelCode)
        cboVehicles.ListIndex = SelectCombo(cboVehicles, Null2String(RS!ModelDescript))
        cboColors.ListIndex = SelectCombo(cboColors, Null2String(RS!Color))
        cboTerms.ListIndex = SelectCombo(cboTerms, Null2String(RS!Terms))
        If IsNull(RS!ExpectedPurchase) = False Then
            txtExpectedPurchase = DateValue(RS!ExpectedPurchase)
        Else
            txtExpectedPurchase = Null
        End If
        txtNotes = Null2String(RS!Notes)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub UpdateLog()
    Dim TSQL                                                          As String
    TSQL = " DECLARE @DT DATETIME " & vbCrLf
    TSQL = TSQL & " SELECT @DT=MAX(StartDateTime) FROM CRIS_SalesAppointments  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
    TSQL = TSQL & " BEGIN " & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGAPPOINTMENT=@DT , HITCOUNTER=1  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End " & vbCrLf
    TSQL = TSQL & " Else " & vbCrLf
    TSQL = TSQL & " BEGIN" & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGAPPOINTMENT=NULL  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End"
    gconDMIS.Execute (TSQL)
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

Friend Sub AddSalesAppointment(xProspectID As Long)
    PROSPECTID = xProspectID
End Sub

Friend Sub EditAppointment(ID As Long, xProspectID As Long)
    AppointmentID = ID
    PROSPECTID = xProspectID
End Sub

Private Sub cboVehicles_Click()
    If cboVehicles.ListIndex = -1 Then Exit Sub
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT CODE , MODEL FROM  ALL_MODEL  WHERE ID=" & cboVehicles.ItemData(cboVehicles.ListIndex))
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        txtModelCode = Null2String(TEMPRS!CODE)
        txtModel = Null2String(TEMPRS!Model)
    End If
End Sub

Private Sub cmdAdd_Click()

    If Function_Access(LOGID, "Acess_ADD", "LOG SALES APPOINTMENT") = False Then Exit Sub
    On Error GoTo Errorcode:

    AppointmentID = 0
    InitMemVars
    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    PICSEARCH.Enabled = False
    On Error Resume Next
    'cboAttendingSE.SetFocus





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    picAdds.Visible = True
    picSaves.Visible = False
    picDataEntry.Enabled = False
    PICSEARCH.Enabled = True
    AppointmentID = 0
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "LOG SALES APPOINTMENT") = False Then Exit Sub
    On Error GoTo Errorcode:

    If ShowConfirmDelete = True Then
        SQL_STATEMENT = "delete from CRIS_SalesAppointments where AppointmentID=" & AppointmentID
        gconDMIS.Execute (SQL_STATEMENT)

        NEW_LogAudit "X", "LOG SALES APPOINTMENT", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""
        UpdateLog
        FillSearchGrid txtSEARCH
        rsRefresh
        StoreMemVars
        If FormExist("MainForm") Then
            MainForm.ShowStatus PROSPECTID
        End If
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()

    If Function_Access(LOGID, "Acess_EDIT", "LOG SALES APPOINTMENT") = False Then Exit Sub
    On Error GoTo Errorcode:

    picAdds.Visible = False
    picSaves.Visible = True
    picDataEntry.Enabled = True
    PICSEARCH.Enabled = False
    On Error Resume Next
    cboAttendingSE.SetFocus





    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    On Error GoTo Errorcode:

    Unload Me





    Exit Sub
Errorcode:
    ShowVBError
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
    Dim SAE                                                           As String
    Dim StartDateTime                                                 As String
    Dim EndDateTime                                                   As String
    Dim Model                                                         As String
    Dim Color                                                         As String
    Dim Terms                                                         As String
    Dim ExpectedPurchase                                              As String
    Dim ModelCode                                                     As String
    Dim ModelDescript                                                 As String
    Dim Note                                                          As String
    Dim SQL                                                           As String

    On Error GoTo Errorcode:

    SAE = N2Str2Null(cboAttendingSE.Text)


    StartDateTime = N2Str2Null(DateValue(txtDate.Value) & " " & TimeValue(txtStartTime))
    EndDateTime = N2Str2Null(DateValue(txtDate.Value) & " " & TimeValue(txtEndTime))
    Model = N2Str2Null(txtModel)
    Color = N2Str2Null(cboColors)
    Terms = N2Str2Null(cboTerms)
    ExpectedPurchase = N2Str2Null(txtExpectedPurchase.Value)
    ModelCode = N2Str2Null(txtModelCode)
    ModelDescript = N2Str2Null(cboVehicles)
    Note = N2Str2Null(txtNotes)

    If AppointmentID <= 0 Then
        SQL = "INSERT INTO CRIS_SalesAppointments("
        SQL = SQL & " ProspectID, SAE, StartDateTime, EndDateTime,  Model, "
        SQL = SQL & " Color, Terms, ExpectedPurchase, ModelCode,ModelDescript,notes) " & vbCrLf
        SQL = SQL & " VALUES("
        SQL = SQL & PROSPECTID & " ,"
        SQL = SQL & SAE & ","
        SQL = SQL & StartDateTime & ","
        SQL = SQL & EndDateTime & ","
        SQL = SQL & Model & ","
        SQL = SQL & Color & ","
        SQL = SQL & Terms & ","
        SQL = SQL & ExpectedPurchase & ","
        SQL = SQL & ModelCode & ","
        SQL = SQL & ModelDescript & ","
        SQL = SQL & Note & ")" & vbCrLf & "SELECT @@IDENTITY"
        
        ' ******* jbf 06282010   **** double entry sa SAE calendar
        'gconDMIS.Execute (SQL)
        'SQL_STATEMENT = SQL
        'NEW_LogAudit "A", "LOG SALES APPOINTMENT", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""
        ' ********************************************************
        
        ' UPDATE: JBF
        ' ************************************************************
        Dim TEMPRS                                                        As ADODB.Recordset
        Set TEMPRS = gconDMIS.Execute(SQL)
        gconDMIS.Execute ("update CRIS_PROSPECTs SET LogAppointment=" & StartDateTime & " where prospectid=" & PROSPECTID)

        
        Set TEMPRS = TEMPRS.NextRecordset
        If Not TEMPRS Is Nothing Then
        AppointmentID = TEMPRS.Collect(0)
        End If
        UpdateLog
        rsRefresh
        RS.Find ("AppointmentID=" & AppointmentID)
        
        NEW_LogAudit "A", "LOG SALES APPOINTMENT", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""
        ' ************************************************************
    Else

        SQL = " Update CRIS_SalesAppointments SET "
        SQL = SQL & " ProspectID=" & PROSPECTID & ", "
        SQL = SQL & " SAE= " & SAE & " ,"
        SQL = SQL & " StartDateTime=" & StartDateTime & ", "
        SQL = SQL & " EndDateTime=" & EndDateTime & ", "
        SQL = SQL & " Model= " & Model & ", "
        SQL = SQL & " ModelCode = " & ModelCode & ", "
        SQL = SQL & " ModelDescript = " & ModelDescript & ", "
        SQL = SQL & " Color=" & Color & ", "
        SQL = SQL & " Terms=" & Terms & ", "
        SQL = SQL & " ExpectedPurchase=" & ExpectedPurchase
        SQL = SQL & " WHERE AppointmentID=" & AppointmentID
        gconDMIS.Execute (SQL)
        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "LOG SALES APPOINTMENT", SQL_STATEMENT, Null2String(PROSPECTID), "", "Prospect ID:" & PROSPECTID, "", ""
    End If

        FillSearchGrid txtSEARCH
        cmdCancel.Value = True
        Set TEMPRS = Nothing
        If FormExist("MainForm") Then
            MainForm.ShowStatus PROSPECTID
        End If

        
    If AppointmentID <= 0 Then
        MessagePop RecSaveOk, "Record Added ", "New Schedule Sucessfully Added", 500, 1
    Else
        MessagePop RecSaveOk, "RecordSaved", "Schedule Sucessfully Updated", 500, 1
    End If



' ******** JBF ***** jun 28,2010 *************
'    Dim TEMPRS                                                        As ADODB.Recordset
'    Set TEMPRS = gconDMIS.Execute(SQL)
'    gconDMIS.Execute ("update CRIS_PROSPECTs SET LogAppointment=" & StartDateTime & " where prospectid=" & PROSPECTID)
'
'    If AppointmentID <= 0 Then
'        MessagePop RecSaveOk, "Record Added ", "New Schedule Sucessfully Added", 500, 1
'    Else
'        MessagePop RecSaveOk, "RecordSaved", "Schedule Sucessfully Updated", 500, 1
'    End If
'
'    Set TEMPRS = TEMPRS.NextRecordset
'    If Not TEMPRS Is Nothing Then
'        AppointmentID = TEMPRS.Collect(0)
'    End If
'    UpdateLog
'    rsRefresh
'    RS.Find ("AppointmentID=" & AppointmentID)
'    FillSearchGrid txtSearch
'    cmdCancel.Value = True
'    Set TEMPRS = Nothing
'    If FormExist("MainForm") Then
'        MainForm.ShowStatus PROSPECTID
'    End If

'**********************************************




    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdViewVStat_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl Is Nothing Then Exit Sub
    If KeyCode = 13 And (Left(ActiveControl.Name, 3) = "txt" Or Left(ActiveControl.Name, 3) = "cbo") Then
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (LOG SALES APPOINTMENT)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(PROSPECTID), "LOG SALES APPOINTMENT")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitMemVars
    InitData
    rsRefresh
    PICSEARCH.Enabled = True: picDataEntry.Enabled = False
    picSaves.Visible = False: picAdds.Visible = True
    StoreMemVars
    SetEntityDetails PROSPECTID, vbNullString

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PROSPECTID = 0
    AppointmentID = 0
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

Private Sub ListView1_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    RS.MoveFirst
    RS.Find ("APPOINTMENTID=" & ITEM.ListSubItems(2).Text)
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

Private Sub txtSearch_Change()
    FillSearchGrid txtSEARCH
End Sub

Private Sub txtStartTime_Change()
    If AppointmentID = 0 Then
        txtEndTime = DateAdd("n", 30, TimeValue(txtStartTime))
    End If

End Sub

Private Sub txtStartTime_LostFocus()
    If AppointmentID = 0 Then
        txtEndTime = DateAdd("n", 30, TimeValue(txtStartTime))
    End If

End Sub

