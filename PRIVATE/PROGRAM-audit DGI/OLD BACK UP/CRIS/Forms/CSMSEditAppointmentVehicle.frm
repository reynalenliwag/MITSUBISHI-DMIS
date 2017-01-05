VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSEditAppVehicle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Vehicle"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7665
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   65
      Top             =   5205
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
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   67
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
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   66
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
      TabIndex        =   60
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
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":11FC
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":134E
         Style           =   1  'Graphical
         TabIndex        =   64
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
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":16B4
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":1806
         Style           =   1  'Graphical
         TabIndex        =   63
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
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":1B31
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":1C83
         Style           =   1  'Graphical
         TabIndex        =   62
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
         MouseIcon       =   "CSMSEditAppointmentVehicle.frx":1FDF
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditAppointmentVehicle.frx":2131
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Add Vehicle"
         Top             =   45
         Width           =   705
      End
   End
   Begin MSComctlLib.ListView lstCons 
      Height          =   1665
      Left            =   90
      TabIndex        =   59
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
      Height          =   4095
      Left            =   120
      TabIndex        =   5
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
         TabIndex        =   68
         Top             =   3390
         Width           =   7815
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   2925
         Left            =   4920
         TabIndex        =   6
         Top             =   150
         Width           =   4035
         Begin VB.TextBox txtprdn 
            Height          =   330
            Left            =   1650
            MaxLength       =   6
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   60
            Width           =   2265
         End
         Begin VB.TextBox txtKMR 
            Height          =   330
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   1140
            Width           =   2265
         End
         Begin VB.TextBox txtVIN 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   20
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   1500
            Width           =   2265
         End
         Begin VB.TextBox txtTIN 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   15
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   420
            Width           =   2265
         End
         Begin VB.TextBox txtWCN 
            Height          =   330
            Left            =   1650
            MaxLength       =   15
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   780
            Width           =   2265
         End
         Begin VB.TextBox txtdateSold 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   2220
            Width           =   2265
         End
         Begin VB.TextBox txtDateDel 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   2580
            Width           =   2265
         End
         Begin VB.TextBox txtConduction 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   8
            TabIndex        =   18
            Text            =   "Text1"
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
            TabIndex        =   7
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
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   10
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
            TabIndex        =   14
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
            TabIndex        =   20
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
            TabIndex        =   22
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
            TabIndex        =   17
            Top             =   1890
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2805
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   4725
         Begin VB.TextBox txtyear 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtMake 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   390
            Width           =   2475
         End
         Begin VB.TextBox txtModel 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   780
            Width           =   2475
         End
         Begin VB.TextBox txtEngine 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   33
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
            MaxLength       =   8
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   1560
            Width           =   2475
         End
         Begin VB.TextBox txtSerial 
            Height          =   330
            Left            =   900
            MaxLength       =   18
            TabIndex        =   40
            Text            =   "Text1"
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
            TabIndex        =   35
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
            TabIndex        =   38
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
            TabIndex        =   39
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   27
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
            TabIndex        =   31
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
            TabIndex        =   34
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
            TabIndex        =   36
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
            TabIndex        =   41
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
            TabIndex        =   42
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
            TabIndex        =   28
            Top             =   120
            Width           =   2205
         End
         Begin VB.Label labID 
            Caption         =   "Label4"
            Height          =   285
            Left            =   3480
            TabIndex        =   30
            Top             =   1260
            Visible         =   0   'False
            Width           =   1185
         End
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
         TabIndex        =   69
         Top             =   3090
         Width           =   2805
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   4770
      ScaleHeight     =   3315
      ScaleWidth      =   4155
      TabIndex        =   43
      Top             =   1650
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
         TabIndex        =   44
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
         TabIndex        =   46
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
         TabIndex        =   48
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
         TabIndex        =   50
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
         TabIndex        =   52
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
         TabIndex        =   54
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
         TabIndex        =   56
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
         TabIndex        =   58
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
         TabIndex        =   45
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
         TabIndex        =   47
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
         TabIndex        =   49
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
         TabIndex        =   51
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
         TabIndex        =   53
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
         TabIndex        =   55
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
         TabIndex        =   57
         Top             =   2400
         Width           =   1485
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
      TabIndex        =   0
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   3
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
      TabIndex        =   4
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
Dim AddorEdit                                                         As String
Dim rsVehicle                                                         As ADODB.Recordset
Public CustomerCode                                                   As String

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
    On Error GoTo Errorcode:
    If MsgBox("Delete this item..." & vbCrLf & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, "Message Box") = vbNo Then
        Exit Sub
    End If
    gconDMIS.Execute "delete  from [CSMS_Cusveh] where id = " & Val(labID.Caption) & ""
    StoreMemVars
    Exit Sub
Errorcode:
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
        Dim rsLoad                                                    As ADODB.Recordset
        Set rsLoad = New ADODB.Recordset
        Set rsLoad = gconDMIS.Execute("Select * from All_Engine where ENGINE = '" & Trim(txtEngine) & "'")
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
    txtyear = ""
    txtMake = ""
    txtModel = ""
    txtEngine = ""
    txtPlateNo = ""
    txtColor = ""
    txtSerial = ""
    txtprdn = ""
    txtTIN = ""
    txtWCN = ""
    txtKMR = ""
    txtVIN = ""
    txtConduction = ""
    labID.Caption = ""
    txtdateSold = ""
    txtDateDel = ""
    txtDescription.Text = ""
    Frame1.Enabled = False
    lstCons.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Sub CheckIfPlateNoAlreadyExist(EXIST As Boolean)
    Dim rsTmp                                                         As New ADODB.Recordset

    Set rsTmp = gconDMIS.Execute("Select Plate_NO From CSMS_CusVeh Where Plate_NO = '" & txtPlateNo.Text & "'")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        EXIST = True
    Else
        EXIST = False
    End If

    Set rsTmp = Nothing
End Sub

Private Sub cmdSave_Click()

    On Error GoTo Errorcode:

    If LTrim(RTrim(txtPlateNo.Text)) = "" Then
        ShowIsRequiredMsg "Plate Number"
        On Error Resume Next
        txtPlateNo.SetFocus
        Exit Sub
    End If

    If LTrim(RTrim(labCustomer.Caption)) = "" Then
        ShowIsRequiredMsg "Customers Name"
        Exit Sub
    End If

    Dim xCuscde, xNIYM, xYear, xMake, xModel, xENGINE, xPLATE_NO, xCLRCDE, xSERIAL, xPRODNO, xTIN_NUMBER, xWAR_CERT, xKMreading, xVIN, xVCOND_NO, xD_SOLD, xDEL_DATE As String
    Dim SQL                             As String

    xCuscde = N2Str2Null(labCustCode)
    xNIYM = N2Str2Null(labCustomer.Caption)
    xYear = N2Str2Null(txtyear)
    xMake = N2Str2Null(txtMake)
    xModel = N2Str2Null(txtModel)
    xENGINE = N2Str2Null(txtEngine)
    xPLATE_NO = N2Str2Null(txtPlateNo)
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
    
    LogAudit "E", "CUSTOMERS VEHICLE ", txtPlateNo
    
    SQL = "UPDATE CSMS_Cusveh SET "
    SQL = SQL & " CUSCDE=" & xCuscde & ","
    SQL = SQL & " NIYM=" & xNIYM & ","
    SQL = SQL & " Yer=" & xYear & ","
    SQL = SQL & " Make=" & xMake & ","
    SQL = SQL & " MODEL=" & xModel & ","
    SQL = SQL & " ENGINE=" & xENGINE & ","
    SQL = SQL & " PLATE_NO=" & xPLATE_NO & ","
    SQL = SQL & " CLRCDE=" & xCLRCDE & ","
    SQL = SQL & " SERIAL=" & xSERIAL & ","
    SQL = SQL & " PRODNO=" & xPRODNO & ","
    SQL = SQL & " TIN_NUMBER=" & xTIN_NUMBER & ","
    SQL = SQL & " WAR_CERT=" & xWAR_CERT & ","
    SQL = SQL & " KMReading=" & xKMreading & ","
    SQL = SQL & " VIN=" & xVIN & ","
    SQL = SQL & " VCOND_NO=" & xVCOND_NO & ","
    SQL = SQL & " D_SOLD=" & xD_SOLD & ","
    SQL = SQL & " Description=" & N2Str2Null(txtDescription) & ","
    SQL = SQL & " DEL_DATE=" & xDEL_DATE
    SQL = SQL & " WHERE ID=" & labID
    gconDMIS.Execute SQL
    frmCSMSEditAppointment.SetVehicleInfo txtPlateNo.Text
    cmdCancel.Value = True

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Function SetColor(CCC As String)
    Dim rsColor                                                       As ADODB.Recordset
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select COLOR_CODE,COLOR_DESC from ALL_Color where COLOR_DESC = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        SetColor = Null2String(rsColor!Color_Code)
    Else
        SetColor = ""
    End If
End Function

Function GetColor(CCC As String)
    Dim rsColor                                                       As ADODB.Recordset
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select COLOR_DESC from ALL_Color where COLOR_CODE = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        GetColor = Null2String(rsColor!Color_Desc)
    Else
        GetColor = ""
    End If
End Function

Private Sub cmdSelect_Click()
With frmCSMSYrMkMlEgn
    .labformname = "frmCSMSEditAppVehicle"
    .Show 1
    
End With
End Sub

Private Sub Command2_Click()
    frmCSMSGetColor.Show 1
End Sub

Private Sub Command3_Click()
    Picture3.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    rsRefresh
    StoreMemVars
End Sub

Sub rsRefresh()
    Set rsVehicle = New ADODB.Recordset
    Set rsVehicle = gconDMIS.Execute("select * from CSMS_Cusveh where PLATE_NO = '" & EDIT_RO & "'")
End Sub

Sub StoreMemVars()
    If Not (rsVehicle.EOF And rsVehicle.BOF) Then
        txtyear = Null2String(rsVehicle!Yer)
        txtMake = Null2String(rsVehicle!Make)
        txtModel = Null2String(rsVehicle!Model)
        txtEngine = Null2String(rsVehicle!Engine)
        txtPlateNo = Null2String(rsVehicle!Plate_no)
        txtColor = GetColor(Null2String(rsVehicle!ClrCde))
        txtSerial = Null2String(rsVehicle!Serial)
        txtprdn = Null2String(rsVehicle!ProdNo)
        txtTIN = Null2String(rsVehicle!TIN_Number)
        txtWCN = Null2String(rsVehicle!War_Cert)
        txtKMR = NumericVal(rsVehicle!KMReading)
        txtVIN = Null2String(rsVehicle!Vin)
        txtConduction = Null2String(rsVehicle!VCond_No)
        txtdateSold = Null2String(rsVehicle!D_Sold)
        txtDateDel = Null2String(rsVehicle!Del_Date)
        txtDescription = Null2String(rsVehicle!Description)
        labID.Caption = rsVehicle!ID
        labVechicle.Caption = Trim(txtyear) & "  " & Trim(txtMake) & "  " & Trim(txtModel) & "  " & Trim(txtEngine)
    Else
        ShowNoRecord
    End If
End Sub

Private Sub txtConduction_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtDateDel_LostFocus()
    If IsDate(txtDateDel) = False Then
        txtDateDel = ""
    End If
End Sub

Private Sub txtdateSold_LostFocus()
    If IsDate(txtdateSold) = False Then
        txtdateSold = ""
    End If
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
