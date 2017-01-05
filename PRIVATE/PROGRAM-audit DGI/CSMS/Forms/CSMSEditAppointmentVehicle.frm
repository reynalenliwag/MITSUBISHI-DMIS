VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSEditAppVehicle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Vehicle"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "CSMSEditAppointmentVehicle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picINS 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   1778
      ScaleHeight     =   3765
      ScaleWidth      =   5565
      TabIndex        =   77
      Top             =   1140
      Visible         =   0   'False
      Width           =   5595
      Begin VB.TextBox txtINSTYPE 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   32
         Top             =   330
         Width           =   1725
      End
      Begin VB.TextBox txtINSCOMP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   33
         Top             =   720
         Width           =   3525
      End
      Begin VB.TextBox txtFINTYPE 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   35
         Top             =   1500
         Width           =   1725
      End
      Begin VB.TextBox txtFINCOMP 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   36
         Top             =   1890
         Width           =   3525
      End
      Begin VB.CommandButton Command1 
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
         Left            =   4740
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Cancel"
         Top             =   2880
         Width           =   705
      End
      Begin MSComCtl2.DTPicker dtpINS 
         Height          =   345
         Left            =   1920
         TabIndex        =   34
         Top             =   1110
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51380225
         CurrentDate     =   39647
      End
      Begin MSComCtl2.DTPicker dtpFIN 
         Height          =   345
         Left            =   1920
         TabIndex        =   37
         Top             =   2280
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51380225
         CurrentDate     =   39647
      End
      Begin VB.CommandButton Command4 
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
         Left            =   4050
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Save New Vehicle"
         Top             =   2880
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Left            =   0
         TabIndex        =   86
         Top             =   0
         Width           =   5565
         _Version        =   655364
         _ExtentX        =   9816
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "FINANCE AND INSURANCE INFO."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16711680
         GradientColorDark=   8388608
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Type"
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
         Index           =   19
         Left            =   540
         TabIndex        =   85
         Top             =   420
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Comapany"
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
         Index           =   21
         Left            =   45
         TabIndex        =   84
         Top             =   840
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration Date"
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
         Index           =   22
         Left            =   555
         TabIndex        =   83
         Top             =   1230
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finance Type"
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
         Index           =   23
         Left            =   735
         TabIndex        =   82
         Top             =   1620
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finance Company"
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
         Index           =   24
         Left            =   345
         TabIndex        =   81
         Top             =   2010
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration Date"
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
         Index           =   25
         Left            =   555
         TabIndex        =   80
         Top             =   2370
         Width           =   1290
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3285
      Left            =   4890
      ScaleHeight     =   3255
      ScaleWidth      =   4155
      TabIndex        =   61
      Top             =   1140
      Visible         =   0   'False
      Width           =   4185
      Begin VB.TextBox txtEnginetype 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1050
         TabIndex        =   24
         Top             =   120
         Width           =   3015
      End
      Begin VB.TextBox txtLiters 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   25
         Top             =   480
         Width           =   1245
      End
      Begin VB.TextBox txtCubic 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   26
         Top             =   840
         Width           =   1245
      End
      Begin VB.TextBox txtDisplacement 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   27
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox txtFuelType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   28
         Top             =   1560
         Width           =   1245
      End
      Begin VB.ComboBox cboAspiration 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         TabIndex        =   29
         Top             =   1950
         Width           =   1515
      End
      Begin VB.TextBox txtEngineVIN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   30
         Top             =   2340
         Width           =   1245
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2670
         TabIndex        =   31
         ToolTipText     =   "Close"
         Top             =   2790
         Width           =   1395
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   62
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Liters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   63
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Cubic Centimeters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   660
         TabIndex        =   64
         Top             =   900
         Width           =   1485
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Cubic Inch Displacement"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   65
         Top             =   1260
         Width           =   1875
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1230
         TabIndex        =   66
         Top             =   1620
         Width           =   1485
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Aspiration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   67
         Top             =   2010
         Width           =   1485
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine VIN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1110
         TabIndex        =   68
         Top             =   2400
         Width           =   1485
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7665
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   75
      Top             =   5415
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
         Left            =   690
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":11FC
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":134E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cancel"
         Top             =   30
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
         Left            =   0
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":168C
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":17DE
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save New Vehicle"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   3000
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   70
      Top             =   8730
      Width           =   6075
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
         Left            =   5355
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":1B2E
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":1C80
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Exit Window"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
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
         Left            =   4665
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":1FE6
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":2138
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Delete Selected Vehicle"
         Top             =   45
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
         Left            =   3975
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":2463
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":25B5
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Edit Selected Vehicle"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   3285
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":2911
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":2A63
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Add Vehicle"
         Top             =   45
         Width           =   705
      End
   End
   Begin MSComctlLib.ListView lstCons 
      Height          =   1665
      Left            =   90
      TabIndex        =   69
      Top             =   7020
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2937
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Make"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Model"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Engine"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Plate No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Color"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Serial No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Prod'n No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "TIN No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Warranty Cert. No."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "KM Reading"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "VIN No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Conductn No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Date Sold"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Date Delivered"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vehicle Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   41
      Top             =   1020
      Width           =   8985
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   21
         Top             =   3570
         Width           =   7815
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   2925
         Left            =   4920
         TabIndex        =   42
         Top             =   150
         Width           =   4035
         Begin VB.TextBox txtprdn 
            Height          =   330
            Left            =   1650
            MaxLength       =   6
            TabIndex        =   12
            Top             =   60
            Width           =   2265
         End
         Begin VB.TextBox txtKMR 
            Height          =   330
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   15
            Top             =   1140
            Width           =   2265
         End
         Begin VB.TextBox txtVIN 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   20
            TabIndex        =   16
            Top             =   1500
            Width           =   2265
         End
         Begin VB.TextBox txtTIN 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   15
            TabIndex        =   13
            Top             =   420
            Width           =   2265
         End
         Begin VB.TextBox txtWCN 
            Height          =   330
            Left            =   1650
            MaxLength       =   15
            TabIndex        =   14
            Top             =   780
            Width           =   2265
         End
         Begin VB.TextBox txtdateSold 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   18
            Top             =   2220
            Width           =   2265
         End
         Begin VB.TextBox txtDateDel 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   19
            Top             =   2580
            Width           =   2265
         End
         Begin VB.TextBox txtConduction 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   8
            TabIndex        =   17
            Top             =   1860
            Width           =   2265
         End
         Begin VB.Label Label1 
            Caption         =   "Production No."
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
            Height          =   255
            Index           =   6
            Left            =   330
            TabIndex        =   43
            Top             =   120
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Kilometer Reading"
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
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   45
            Top             =   1170
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   "VIN No."
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
            Height          =   255
            Index           =   8
            Left            =   930
            TabIndex        =   46
            Top             =   1530
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "TIN No."
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
            Height          =   255
            Index           =   9
            Left            =   930
            TabIndex        =   44
            Top             =   450
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Warranty Cert. No."
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
            Height          =   255
            Index           =   10
            Left            =   60
            TabIndex        =   47
            Top             =   810
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   "Date Sold"
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
            Height          =   255
            Index           =   11
            Left            =   690
            TabIndex        =   49
            Top             =   2250
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   "Date Delivered"
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
            Height          =   255
            Index           =   12
            Left            =   330
            TabIndex        =   50
            Top             =   2610
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Conduction No."
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
            Height          =   255
            Index           =   14
            Left            =   270
            TabIndex        =   48
            Top             =   1890
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2805
         Left            =   120
         TabIndex        =   51
         Top             =   300
         Width           =   4725
         Begin VB.TextBox txtyear 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtMake 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   390
            Width           =   2475
         End
         Begin VB.TextBox txtModel 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   780
            Width           =   2475
         End
         Begin VB.TextBox txtEngine 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1170
            Width           =   2475
         End
         Begin VB.TextBox txtPlateno 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   900
            MaxLength       =   7
            TabIndex        =   8
            Top             =   1560
            Width           =   2475
         End
         Begin VB.TextBox txtSerial 
            Height          =   330
            Left            =   900
            MaxLength       =   18
            TabIndex        =   11
            Top             =   2340
            Width           =   2475
         End
         Begin VB.CommandButton cmdDetails 
            Caption         =   "&Details..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3420
            TabIndex        =   7
            ToolTipText     =   "Details"
            Top             =   1170
            Width           =   765
         End
         Begin VB.TextBox txtColor 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1950
            Width           =   2475
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
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
            Left            =   3420
            TabIndex        =   10
            Top             =   1950
            Width           =   345
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Select Year, Make and Model"
            Top             =   0
            Width           =   345
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   0
            Left            =   420
            TabIndex        =   52
            Top             =   60
            Width           =   495
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   53
            Top             =   450
            Width           =   525
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   2
            Left            =   300
            TabIndex        =   56
            Top             =   840
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Engine"
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
            Height          =   255
            Index           =   3
            Left            =   210
            TabIndex        =   57
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Plate No."
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
            Height          =   255
            Index           =   4
            Left            =   60
            TabIndex        =   58
            Top             =   1590
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Serial No."
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
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   59
            Top             =   2370
            Width           =   915
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   13
            Left            =   360
            TabIndex        =   60
            Top             =   1980
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Select  Year, Make and Model"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   15
            Left            =   2340
            TabIndex        =   54
            Top             =   120
            Width           =   2205
         End
         Begin VB.Label labID 
            Caption         =   "Label4"
            Height          =   285
            Left            =   3480
            TabIndex        =   55
            Top             =   1260
            Visible         =   0   'False
            Width           =   1185
         End
      End
      Begin wizButton.cmd cmdINS 
         Height          =   405
         Left            =   5280
         TabIndex        =   20
         Top             =   3120
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   714
         TX              =   "Edit Finance and Insurance info."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "CSMSEditAppointmentVehicle.frx":2D76
      End
      Begin VB.Label Label1 
         Caption         =   "Customer Vehicle Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   76
         Top             =   3270
         Width           =   2805
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Code"
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
      Left            =   210
      TabIndex        =   38
      Top             =   120
      Width           =   1365
   End
   Begin VB.Label labCustomer 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1890
      TabIndex        =   1
      Top             =   540
      Width           =   7215
   End
   Begin VB.Label labCustCode 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1890
      TabIndex        =   0
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Customer Name"
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
      Left            =   210
      TabIndex        =   39
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label labVechicle 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
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
      Left            =   1890
      TabIndex        =   40
      Top             =   720
      Width           =   5115
   End
End
Attribute VB_Name = "frmCSMSEditAppVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                          As String
Dim rsVehicle                                          As ADODB.Recordset
Public CustomerCode                                    As String
Dim WithEvents frmSelectMakeMode                       As frmCSMSYrMkMlEgn
Attribute frmSelectMakeMode.VB_VarHelpID = -1


Private Sub frmSelectMakeMode_SelectedDetails(XYEAR As String, xMake As String, xModel As String, xENGINE As String, xModelDescription As String)
    txtYear = XYEAR
    TXTMAKE = xMake
    txtModel = xModel
    txtENGINE = xENGINE
    txtDescription = xModelDescription

End Sub


Function CheckIfPlateNoAlreadyExist(vFIELD As String, VKEY As String, VTYPE As Integer, VFNAME As String) As Boolean
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select " & vFIELD & ",NIYM From CSMS_CusVeh Where " & vFIELD & " = '" & VKEY & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If VTYPE = 1 Then MsgBox "" & VFNAME & " " & VKEY & " already exist and registered to " & RSTMP!NIYM & "", vbInformation, "CSMS"
        CheckIfPlateNoAlreadyExist = True
    Else
        CheckIfPlateNoAlreadyExist = False
    End If

    Set RSTMP = Nothing
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

Function GetColor(CCC As String)
    Dim rsColor                                        As ADODB.Recordset
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select COLOR_DESC from ALL_Color where COLOR_CODE = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        GetColor = Null2String(rsColor!color_desc)
    Else
        GetColor = ""
    End If
End Function

Sub rsRefresh()
    Set rsVehicle = New ADODB.Recordset
    Set rsVehicle = gconDMIS.Execute("select * from CSMS_Cusveh where PLATE_NO = '" & EDIT_RO & "'")
End Sub

Sub StoreMemVars()
    If Not (rsVehicle.EOF And rsVehicle.BOF) Then
        txtYear = Null2String(rsVehicle!YER)
        TXTMAKE = Null2String(rsVehicle!Make)
        txtModel = Null2String(rsVehicle!MODEL)
        txtENGINE = Null2String(rsVehicle!Engine)
        txtPLATENO = Null2String(rsVehicle!PLATE_NO)
        txtColor = GetColor(Null2String(rsVehicle!ClrCde))
        txtSerial = Null2String(rsVehicle!SERIAL)
        txtprdn = Null2String(rsVehicle!ProdNo)
        txtTIN = Null2String(rsVehicle!TIN_Number)
        txtWCN = Null2String(rsVehicle!War_Cert)
        txtKMR = NumericVal(rsVehicle!KMReading)
        txtVIN = Null2String(rsVehicle!Vin)
        txtConduction = Null2String(rsVehicle!VCOND_NO)
        txtdateSold = Null2String(rsVehicle!D_SOLD)
        txtDateDel = Null2String(rsVehicle!DEL_DATE)
        txtDescription = Null2String(rsVehicle!Description)
        labid.Caption = rsVehicle!ID
        labVechicle.Caption = Trim(txtYear) & "  " & Trim(TXTMAKE) & "  " & Trim(txtModel) & "  " & Trim(txtENGINE)

        'UPDATE BY: MJP 07182008 6:00 PM
        txtINSTYPE.Text = Null2String(rsVehicle!INS_TYPE)
        txtINSCOMP.Text = Null2String(rsVehicle!INS_COMP)
        If Null2String(rsVehicle!INS_EXP_DATE) = "" Then
            dtpINS.Value = Date
        Else
            dtpINS.Value = Null2String(rsVehicle!INS_EXP_DATE)
        End If

        txtFINTYPE.Text = Null2String(rsVehicle!FIN_TYPE)
        txtFINCOMP.Text = Null2String(rsVehicle!FIN_COMP)
        If Null2String(rsVehicle!FIN_EXP_DATE) = "" Then
            dtpFIN.Value = Date
        Else
            dtpFIN.Value = Null2String(rsVehicle!FIN_EXP_DATE)
        End If
        'UPDATE BY: MJP 07182008 6:00 PM
    Else
        ShowNoRecord
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "CUSTOMER VEHICLE") = False Then Exit Sub
    AddorEdit = "ADD"
    initMemvars
    picAdds.Visible = False
    picSaves.Visible = True
    Frame1.Enabled = True
    lstCons.Enabled = False
    Me.Caption = "Add Vehicle"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCustomer_Click()
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "CUSTOMER VEHICLE") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox("Delete this item..." & vbCrLf & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    gconDMIS.Execute "delete  from [CSMS_Cusveh] where id = " & Val(labid.Caption) & ""
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdDetails_Click()
    If Picture3.Visible = True Then
        cmdDetails.Caption = "&Close..."
        Picture3.Visible = False
        Picture3.ZOrder 1
    Else
        Picture3.Visible = True
        Picture3.ZOrder 0
        cmdDetails.Caption = "&Details..."
        Dim rsLoad                                     As ADODB.Recordset
        Set rsLoad = New ADODB.Recordset
        Set rsLoad = gconDMIS.Execute("Select * from All_Engine where ENGINE = '" & Trim(txtENGINE) & "'")
        If Not rsLoad.EOF And Not rsLoad.BOF Then
            txtEnginetype = Null2String(rsLoad![Engine])
            txtLiters = Null2String(rsLoad![Liters])
            txtCubic = Null2String(rsLoad![Cubic])
            txtDisplacement = Null2String(rsLoad![Displacement])
            txtFuelType = Null2String(rsLoad![FuelType])
            cboAspiration.Text = Null2String(rsLoad![Aspiration])
            txtEngineVIN = Null2String(rsLoad![EngineVIN])
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "CUSTOMER VEHICLE") = False Then Exit Sub
    AddorEdit = "EDIT"
    picAdds.Visible = False
    picSaves.Visible = True
    Frame1.Enabled = True
    lstCons.Enabled = False

End Sub

Private Sub initMemvars()
    labVechicle.Caption = ""
    txtYear = ""
    TXTMAKE = ""
    txtModel = ""
    txtENGINE = ""
    txtPLATENO = ""
    txtColor = ""
    txtSerial = ""
    txtprdn = ""
    txtTIN = ""
    txtWCN = ""
    txtKMR = ""
    txtVIN = ""
    txtConduction = ""
    labid.Caption = ""
    txtdateSold = ""
    txtDateDel = ""
    txtDescription.Text = ""
    Frame1.Enabled = False
    lstCons.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'UPDATE BY: MJP 07182008 6:00 PM
Private Sub cmdINS_Click()
    Frame1.Enabled = False
    picSaves.Enabled = False
    picINS.Visible = True
    txtINSTYPE.SetFocus
End Sub
'UPDATE BY: MJP 07182008 6:00 PM

Function ReturnVehicleID(VKEY As String, vFIELD As String, VEHID As Integer, KKEY As String) As Boolean
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("select * from csms_Cusveh where " & vFIELD & " = '" & VKEY & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If VEHID = RSTMP!ID Then
            ReturnVehicleID = False
        Else
            MsgBox "" & KKEY & " already exist, Registered to" & vbCrLf & GetAcctName(Null2String(RSTMP!CUSCDE)) & "", vbExclamation, "CSMS"
            ReturnVehicleID = True
        End If
    End If
    Set RSTMP = Nothing
End Function

Function GetAcctName(ACCTNO As String) As String
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE CUSCDE = '" & ACCTNO & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetAcctName = Null2String(RSTMP!ACCTNAME)
    End If
    Set RSTMP = Nothing
End Function

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode:

    If LTrim(RTrim(txtPLATENO.Text)) = "" Then
        ShowIsRequiredMsg "Plate Number"
        On Error Resume Next
        txtPLATENO.SetFocus
        Exit Sub
    End If

    If LTrim(RTrim(labCustomer.Caption)) = "" Then
        ShowIsRequiredMsg "Customers Name"
        Exit Sub
    End If
    If Not txtdateSold.Text = "" Then
        If IsDate(txtdateSold) = False Then
            MsgBox "Date Format is Incorrect", vbCritical, "ERROR"
            txtdateSold.SetFocus
            Exit Sub
        End If
    End If

    If Not txtDateDel.Text = "" Then
        If IsDate(txtDateDel) = False Then
            MsgBox "Date Format is Incorrect", vbCritical, "ERROR"
            txtDateDel.SetFocus
            Exit Sub
        End If
    End If
    Dim xCuscde, xNIYM, XYEAR, xMake, xModel, xENGINE, xPLATE_NO, xCLRCDE, xSERIAL, xPRODNO, xTIN_NUMBER, xWAR_CERT, xKMreading, xVIN, xVCOND_NO, xD_SOLD, xDEL_DATE As String
    Dim SQL                                            As String
    Dim TMP_PLATE                                      As String

    TMP_PLATE = txtPLATENO.Text
    If AddorEdit = "ADD" Then
        If CheckIfPlateNoAlreadyExist("PLATE_NO", TMP_PLATE, 1, "Plate no.") = True Then
            txtPLATENO.SetFocus
            Exit Sub
        End If
    Else
        'UPDATE BY   : MJP 08292008 †-----------------------------------------------------
        'DESCRIPTION : TO LIMIT THE DUPLICATE OF PLATE NO
        If CheckIfPlateNoAlreadyExist("PLATE_NO", TMP_PLATE, 0, "Plate no.") = True Then
            If ReturnVehicleID(TMP_PLATE, "PLATE_NO", labid.Caption, "Plate no.") = True Then
                txtPLATENO.SetFocus
                Exit Sub
            End If
        End If
        'UPDATE BY   : MJP 08292008 ------------------------------------------------------†
        'If CheckIfPlateNoAlreadyExist Then
        '    If Not labID.Caption = ReturnVehicleID(TMP_PLATE) Then
        '        MsgBox "Plate No. Already Exist", vbExclamation, "Add Customer Vehicle"
        '        txtPlateno.SetFocus
        '        Exit Sub
        '    End If
        'End If
    End If

    xCuscde = N2Str2Null(labCustCode)
    xNIYM = N2Str2Null(labCustomer.Caption)
    XYEAR = N2Str2Null(txtYear)
    xMake = N2Str2Null(TXTMAKE)
    xModel = N2Str2Null(txtModel)
    xENGINE = N2Str2Null(txtENGINE)
    xPLATE_NO = N2Str2Null(txtPLATENO)
    xCLRCDE = N2Str2Null(SetColor(txtColor))
    xSERIAL = N2Str2Null(txtSerial)
    xPRODNO = N2Str2Null(txtprdn)
    xTIN_NUMBER = N2Str2Null(txtTIN)
    xWAR_CERT = N2Str2Null(txtWCN)
    xKMreading = N2Str2Null(txtKMR)
    xVIN = N2Str2Null(txtVIN)
    xVCOND_NO = N2Str2Null(txtConduction)
    xD_SOLD = N2Str2Null(txtdateSold)
    xDEL_DATE = N2Str2Null(txtDateDel)

    'UPDATE BY : MJP 08-29-2008 ----------------------------------------------------------
    Dim getOldPlate                                    As New ADODB.Recordset
    Dim RSTMP                                          As New ADODB.Recordset
    Dim vOldPlate                                      As String

    Set getOldPlate = gconDMIS.Execute("select plate_no from csms_cusveh where id = " & labid & "")
    If Not (getOldPlate.BOF And getOldPlate.EOF) Then
        vOldPlate = Null2String(getOldPlate!PLATE_NO)
    End If
    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE PLATE_NO = '" & vOldPlate & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If MsgBox("This vehicle had a past transaction. " & vbCrLf & "editing this information will upadte all past records", vbQuestion + vbYesNo, "Are you sure") = vbNo Then
            GoTo CONT
        Else
            SQL_STATEMENT = "update csms_repor set Plate_no = " & xPLATE_NO & " where plate_no ='" & vOldPlate & "'"
            gconDMIS.Execute SQL_STATEMENT

            'NEW LOG AUDIT-----------------------------------------------------
            Dim rsREPOR                                As New ADODB.Recordset
            Set rsREPOR = gconDMIS.Execute("SELECT REP_OR FROM CSMS_REPOR WHERE PLATE_NO = " & xPLATE_NO & "")
            If Not (rsREPOR.BOF And rsREPOR.EOF) Then
                Do While Not rsREPOR.EOF
                    Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(Null2String(rsREPOR!rep_OR)), "REP_OR", "CSMS_REPOR"), "", "PLATE NO: " & Null2String(xPLATE_NO), "", "")
                    rsREPOR.MoveNext
                Loop
            End If
            'NEW LOG AUDIT-----------------------------------------------------

            gconDMIS.Execute "update csms_repairorder set Plate_no = " & xPLATE_NO & " where plate_no ='" & vOldPlate & "'"
            gconDMIS.Execute "update csms_appointment set plate_no = " & xPLATE_NO & " where plate_no = '" & vOldPlate & "'"
        End If
    End If
    'UPDATE BY : MJP 08-29-2008 ----------------------------------------------------------

    SQL = "UPDATE CSMS_Cusveh SET "
    SQL = SQL & " CUSCDE = " & xCuscde & ","
    SQL = SQL & " NIYM = " & xNIYM & ","
    SQL = SQL & " Yer = " & XYEAR & ","
    SQL = SQL & " Make = " & xMake & ","
    SQL = SQL & " MODEL = " & xModel & ","
    SQL = SQL & " ENGINE = " & xENGINE & ","
    SQL = SQL & " PLATE_NO = " & xPLATE_NO & ","
    SQL = SQL & " CLRCDE = " & xCLRCDE & ","
    SQL = SQL & " SERIAL = " & xSERIAL & ","
    SQL = SQL & " PRODNO = " & xPRODNO & ","
    SQL = SQL & " TIN_NUMBER = " & xTIN_NUMBER & ","
    SQL = SQL & " WAR_CERT = " & xWAR_CERT & ","
    SQL = SQL & " KMReading = " & xKMreading & ","
    SQL = SQL & " VIN = " & xVIN & ","
    SQL = SQL & " VCOND_NO = " & xVCOND_NO & ","
    SQL = SQL & " D_SOLD = " & xD_SOLD & ","
    SQL = SQL & " Description = " & N2Str2Null(txtDescription) & ","
    SQL = SQL & " DEL_DATE = " & xDEL_DATE
    SQL = SQL & " WHERE ID = " & labid
    gconDMIS.Execute SQL

    SQL_STATEMENT = SQL

    'LogAudit "E", "EDIT CUSTOMER INFORMATION", labCustCode & " PLATE#" & txtPlateNo
    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("E", "CUSTOMER VEHICLE", SQL_STATEMENT, labid, "", "PLATE NO: " & Null2String(xPLATE_NO), "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    MessagePop InfoFriend, "Vehicle Information Updated", "Vehicle Information Sucessfully Updated!", 1000
    frmCSMSEditAppointment.SetVehicleInfo txtPLATENO.Text

CONT:
    cmdCancel.Value = True

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSelect_Click()
    Set frmSelectMakeMode = New frmCSMSYrMkMlEgn
    frmSelectMakeMode.Show 1

End Sub

'UPDATE BY: MJP 07182008 6:00 PM
Private Sub Command1_Click()
    Frame1.Enabled = True
    picSaves.Enabled = True
    picINS.Visible = False
End Sub
'UPDATE BY: MJP 07182008 6:00 PM

Private Sub Command2_Click()
    frmCSMSGetColor.Show 1
End Sub

Private Sub Command3_Click()
    Picture3.Visible = False
End Sub

'UPDATE BY: MJP 07182008 6:00 PM
Private Sub Command4_Click()
    If txtINSTYPE.Text = "" Then
        ShowIsRequiredMsg ("Insurance Type")
        txtINSTYPE.SetFocus
        Exit Sub
    End If
    If txtINSCOMP.Text = "" Then
        ShowIsRequiredMsg ("Insurance Company")
        txtINSCOMP.SetFocus
        Exit Sub
    End If
    If txtFINTYPE.Text = "" Then
        ShowIsRequiredMsg ("Finance Type")
        txtFINTYPE.SetFocus
        Exit Sub
    End If
    If txtFINCOMP.Text = "" Then
        ShowIsRequiredMsg ("Finance Company")
        txtFINCOMP.SetFocus
        Exit Sub
    End If

    Dim vINSTYPE                                       As String
    Dim vINSCOMP                                       As String
    Dim VINSDATE                                       As String
    Dim vFINTYPE                                       As String
    Dim vFINCOMP                                       As String
    Dim vFINDATE                                       As String

    vINSTYPE = N2Str2Null(txtINSTYPE)
    vINSCOMP = N2Str2Null(txtINSCOMP)
    VINSDATE = N2Str2Null(dtpINS.Value)
    vFINTYPE = N2Str2Null(txtINSTYPE)
    vFINCOMP = N2Str2Null(txtFINCOMP)
    vFINDATE = N2Str2Null(dtpFIN.Value)

    If MsgBox("Save Insurance and Finance Info.", vbQuestion + vbYesNo, "CSMS") = vbYes Then
        gconDMIS.Execute ("UPDATE CSMS_CUSVEH SET INS_TYPE = " & vINSTYPE & ",INS_COMP = " & vINSCOMP & _
                          ",INS_EXP_DATE = " & VINSDATE & ",FIN_TYPE = " & vFINTYPE & ",FIN_COMP = " & vFINCOMP & _
                          ",FIN_EXP_DATE = " & vFINDATE & " WHERE ID = " & labid.Caption & "")

        Call ShowSuccessFullyUpdated
    End If

    Call Command1_Click
End Sub
'UPDATE BY: MJP 07182008 6:00 PM

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    StoreMemVars
End Sub

Private Sub txtConduction_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtDateDel_LostFocus()
    '    If IsDate(txtDateDel) = False Then
    '        txtDateDel = ""
    '    End If
End Sub

Private Sub txtdateSold_LostFocus()
    '    If IsDate(txtdateSold) = False Then
    '        txtdateSold = ""
    '    End If
End Sub

Private Sub txtKMR_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtPlateno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0: Exit Sub
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtprdn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtTIN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtVIN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtWCN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

