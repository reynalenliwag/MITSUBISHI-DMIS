VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSAddVehicle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Vehicle"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
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
   Icon            =   "CSMSAddVehicle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rptVehicle 
      Left            =   1650
      Top             =   7950
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ListView lstCons 
      Height          =   1725
      Left            =   90
      TabIndex        =   23
      Top             =   6120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3043
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
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   3030
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   68
      Top             =   7860
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
         Left            =   5370
         MouseIcon       =   "CSMSAddVehicle.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Left            =   4680
         MouseIcon       =   "CSMSAddVehicle.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Print this Record"
         Top             =   60
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
         Left            =   3990
         MouseIcon       =   "CSMSAddVehicle.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Delete Selected Vehicle"
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
         Left            =   3300
         MouseIcon       =   "CSMSAddVehicle.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Edit Selected Vehicle"
         Top             =   60
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
         Left            =   2610
         MouseIcon       =   "CSMSAddVehicle.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Add Vehicle"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7695
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   69
      Top             =   7875
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
         MouseIcon       =   "CSMSAddVehicle.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   71
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
         MouseIcon       =   "CSMSAddVehicle.frx":245A
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":25AC
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Save New Vehicle"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vehicle Information"
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
      ForeColor       =   &H00800000&
      Height          =   5055
      Left            =   120
      TabIndex        =   32
      Top             =   1020
      Width           =   8985
      Begin VB.ComboBox cboSelling_Dealer 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1020
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   3540
         Width           =   5535
      End
      Begin VB.ComboBox cboEndUser 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1020
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   3150
         Width           =   5535
      End
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
         Height          =   675
         Left            =   1020
         MaxLength       =   100
         TabIndex        =   22
         Top             =   4260
         Width           =   7815
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2805
         Left            =   120
         TabIndex        =   42
         Top             =   300
         Width           =   4725
         Begin VB.TextBox txtyear 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtMake 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   390
            Width           =   2475
         End
         Begin VB.TextBox txtModel 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   780
            Width           =   2475
         End
         Begin VB.TextBox txtEngine 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   900
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
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
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
            Height          =   345
            Left            =   1920
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Select Year, Make and Model"
            Top             =   0
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   450
            TabIndex        =   43
            Top             =   60
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   1
            Left            =   375
            TabIndex        =   44
            Top             =   450
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   2
            Left            =   330
            TabIndex        =   47
            Top             =   900
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   3
            Left            =   270
            TabIndex        =   48
            Top             =   1260
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   4
            Left            =   90
            TabIndex        =   49
            Top             =   1650
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   5
            Left            =   30
            TabIndex        =   50
            Top             =   2430
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Index           =   13
            Left            =   390
            TabIndex        =   51
            Top             =   2040
            Width           =   450
         End
         Begin VB.Label labID 
            Caption         =   "Label4"
            Height          =   285
            Left            =   3450
            TabIndex        =   46
            Top             =   420
            Visible         =   0   'False
            Width           =   1185
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
            TabIndex        =   45
            Top             =   120
            Width           =   2205
         End
      End
      Begin wizButton.cmd cmdINS 
         Height          =   765
         Left            =   6570
         TabIndex        =   76
         Top             =   3120
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1349
         TX              =   "Edit Finance and Insurance info."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "CSMSAddVehicle.frx":28FC
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   2925
         Left            =   4860
         TabIndex        =   33
         Top             =   180
         Width           =   4065
         Begin VB.TextBox txtDateDel_ 
            BackColor       =   &H000000FF&
            Height          =   330
            Left            =   3990
            TabIndex        =   21
            Top             =   2520
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.TextBox txtdateSold_ 
            BackColor       =   &H000000FF&
            Height          =   330
            Left            =   3990
            TabIndex        =   20
            Top             =   2190
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.TextBox txtprdn 
            Height          =   330
            Left            =   1740
            MaxLength       =   6
            TabIndex        =   14
            Top             =   60
            Width           =   2265
         End
         Begin VB.TextBox txtKMR 
            Height          =   330
            Left            =   1740
            MaxLength       =   10
            TabIndex        =   17
            Top             =   1140
            Width           =   2265
         End
         Begin VB.TextBox txtVIN 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1740
            MaxLength       =   20
            TabIndex        =   18
            Top             =   1500
            Width           =   2265
         End
         Begin VB.TextBox txtTIN 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1740
            MaxLength       =   15
            TabIndex        =   15
            Top             =   420
            Width           =   2265
         End
         Begin VB.TextBox txtWCN 
            Height          =   330
            Left            =   1740
            MaxLength       =   15
            TabIndex        =   16
            Top             =   780
            Width           =   2265
         End
         Begin VB.TextBox txtConduction 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1740
            MaxLength       =   8
            TabIndex        =   19
            Top             =   1860
            Width           =   2265
         End
         Begin MSComCtl2.DTPicker txtdateSold 
            Height          =   315
            Left            =   1740
            TabIndex        =   93
            Top             =   2220
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   54657025
            CurrentDate     =   40091
         End
         Begin MSComCtl2.DTPicker txtDateDel 
            Height          =   315
            Left            =   1740
            TabIndex        =   94
            Top             =   2550
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   54657025
            CurrentDate     =   40091
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   6
            Left            =   450
            TabIndex        =   34
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   7
            Left            =   135
            TabIndex        =   36
            Top             =   1170
            Width           =   1560
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "VIN."
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
            Index           =   8
            Left            =   1365
            TabIndex        =   37
            Top             =   1530
            Width           =   330
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TIN."
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
            Index           =   9
            Left            =   1380
            TabIndex        =   35
            Top             =   450
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   10
            Left            =   135
            TabIndex        =   38
            Top             =   810
            Width           =   1560
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   11
            Left            =   885
            TabIndex        =   40
            Top             =   2250
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   12
            Left            =   465
            TabIndex        =   41
            Top             =   2610
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   14
            Left            =   420
            TabIndex        =   39
            Top             =   1890
            Width           =   1275
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Source"
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
         Index           =   18
         Left            =   345
         TabIndex        =   59
         Top             =   3570
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "End-User"
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
         Index           =   17
         Left            =   180
         TabIndex        =   73
         Top             =   3180
         Width           =   780
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   150
         TabIndex        =   72
         Top             =   3990
         Width           =   2805
      End
   End
   Begin VB.PictureBox picINS 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   1808
      ScaleHeight     =   3765
      ScaleWidth      =   5565
      TabIndex        =   77
      Top             =   2595
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
         TabIndex        =   78
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
         TabIndex        =   79
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
         TabIndex        =   81
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
         TabIndex        =   82
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
         MouseIcon       =   "CSMSAddVehicle.frx":2918
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":2A6A
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Cancel"
         Top             =   2880
         Width           =   705
      End
      Begin MSComCtl2.DTPicker dtpINS 
         Height          =   345
         Left            =   1920
         TabIndex        =   80
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
         Format          =   54657025
         CurrentDate     =   39647
      End
      Begin MSComCtl2.DTPicker dtpFIN 
         Height          =   345
         Left            =   1920
         TabIndex        =   83
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
         Format          =   54657025
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
         MouseIcon       =   "CSMSAddVehicle.frx":2DA8
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":2EFA
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Save New Vehicle"
         Top             =   2880
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Left            =   0
         TabIndex        =   92
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   19
         Left            =   540
         TabIndex        =   91
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   21
         Left            =   45
         TabIndex        =   90
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   22
         Left            =   555
         TabIndex        =   89
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   23
         Left            =   735
         TabIndex        =   88
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   24
         Left            =   345
         TabIndex        =   87
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
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   25
         Left            =   555
         TabIndex        =   86
         Top             =   2370
         Width           =   1290
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   4440
      ScaleHeight     =   3315
      ScaleWidth      =   4155
      TabIndex        =   52
      Top             =   2550
      Visible         =   0   'False
      Width           =   4185
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
         Left            =   2550
         TabIndex        =   74
         Top             =   1200
         Width           =   1245
      End
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
         Left            =   1260
         TabIndex        =   53
         Top             =   120
         Width           =   2535
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
         Left            =   2550
         TabIndex        =   55
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
         Left            =   2550
         TabIndex        =   57
         Top             =   840
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
         Left            =   2550
         TabIndex        =   61
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
         Left            =   2550
         TabIndex        =   63
         Top             =   1950
         Width           =   1305
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
         Left            =   2550
         TabIndex        =   65
         Top             =   2340
         Width           =   1245
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2550
         TabIndex        =   67
         ToolTipText     =   "Close"
         Top             =   2790
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Engine Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   54
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Liters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1725
         TabIndex        =   56
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cubic Centimeters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   660
         TabIndex        =   58
         Top             =   900
         Width           =   1560
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cubic Inch Displacement"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   60
         Top             =   1260
         Width           =   2025
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1425
         TabIndex        =   62
         Top             =   1620
         Width           =   795
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aspiration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1365
         TabIndex        =   64
         Top             =   2010
         Width           =   855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Engine VIN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1350
         TabIndex        =   66
         Top             =   2400
         Width           =   870
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
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
      Index           =   20
      Left            =   7230
      TabIndex        =   75
      Top             =   150
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   210
      TabIndex        =   29
      Top             =   180
      Width           =   1455
   End
   Begin VB.Label labCustomer 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1890
      TabIndex        =   1
      Top             =   540
      Width           =   7215
   End
   Begin VB.Label labCustCode 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1890
      TabIndex        =   0
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   210
      TabIndex        =   30
      Top             =   660
      Width           =   1515
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
      Left            =   3960
      TabIndex        =   31
      Top             =   30
      Visible         =   0   'False
      Width           =   5115
   End
End
Attribute VB_Name = "frmCSMSAddVehicle"
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

Private Sub frmSelectMakeMode_SelectedDetails(XYEAR As String, xMake As String, xMODEL As String, xENGINE As String, xModelDescription As String)
    txtyear = XYEAR
    txtMake = xMake
    txtModel = xMODEL
    txtEngine = xENGINE
    txtDescription = xModelDescription
End Sub

Function CheckIfPlateNoAlreadyExist(vFIELD As String, VKEY As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("Select " & vFIELD & " From CSMS_CusVeh Where " & vFIELD & " = '" & VKEY & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        CheckIfPlateNoAlreadyExist = True
    Else
        CheckIfPlateNoAlreadyExist = False
    End If

    Set rstmp = Nothing
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

Function SetColor(CCC As String)
    Dim rsColor                                        As New ADODB.Recordset
    rsColor.Open "select COLOR_CODE,COLOR_DESC from ALL_Color where COLOR_DESC = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        SetColor = Null2String(rsColor!Color_code)
    Else
        SetColor = ""
    End If
End Function

Function GetColor(CCC As String)
    Dim rsColor                                        As New ADODB.Recordset
    rsColor.Open "select COLOR_DESC from ALL_Color where COLOR_CODE = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        GetColor = Null2String(rsColor!color_desc)
    Else
        GetColor = ""
    End If
End Function

Function SetSellingDealer(XXX As String, CodeOrName As Integer) As String
    Dim rsSellingDealer                                As New ADODB.Recordset
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

Function SetEndUser(XXX As String, CodeOrName As Integer) As String
    Dim rsEndUser                                      As New ADODB.Recordset
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

Sub rsRefresh()
    Set rsVehicle = New ADODB.Recordset
    rsVehicle.Open "select * from CSMS_Cusveh where (CUSCDE = '" & CustomerCode & "' OR ENDUSER = '" & CustomerCode & "')", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not (rsVehicle.EOF And rsVehicle.BOF) Then
        txtyear = Null2String(rsVehicle!YER)
        txtMake = Null2String(rsVehicle!Make)
        txtModel = Null2String(rsVehicle!Model)
        txtEngine = Null2String(rsVehicle!Engine)
        txtPlateno = Null2String(rsVehicle!PLATE_NO)
        txtColor = GetColor(Null2String(rsVehicle!ClrCde))
        txtSerial = Null2String(rsVehicle!serial)
        txtprdn = Null2String(rsVehicle!prodno)
        txtTIN = Null2String(rsVehicle!TIN_Number)
        txtWCN = Null2String(rsVehicle!War_Cert)
        txtKMR = NumericVal(rsVehicle!KMReading)
        txtVIN = Null2String(rsVehicle!Vin)
        txtConduction = Null2String(rsVehicle!VCOND_NO)
        txtdateSold = Null2String(rsVehicle!D_SOLD)
        txtDateDel = Null2String(rsVehicle!DEL_DATE)
        txtDescription = Null2String(rsVehicle!Description)
        labID.Caption = rsVehicle!ID
        labVechicle.Caption = Trim(txtyear) & "  " & Trim(txtMake) & "  " & Trim(txtModel) & "  " & Trim(txtEngine)
        cboEndUser.Text = SetEndUser(Null2String(rsVehicle!ENDUSER), 1)
        cboSelling_Dealer.Text = SetSellingDealer(Null2String(rsVehicle!Selling_Dealer), 1)

        'UPDATE BY: MJP 07182008 6:00 PM
        cmdINS.Visible = True
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
        cmdAdd.Value = True
    End If
End Sub

Sub FillSearchGrid()
    Dim temprs                                         As New ADODB.Recordset

    lstCons.Enabled = False

    Set temprs = gconDMIS.Execute("select YER,MAKE,MODEL,ENGINE,PLATE_NO,CLRCDE,SERIAL,PRODNO,TIN_NUMBER,WAR_CERT,KMreading,VIN,VCOND_NO,D_SOLD,DEL_DATE,ID from CSMS_Cusveh where (CUSCDE = '" & Repleys(CustomerCode) & "' OR ENDUSER = '" & Repleys(CustomerCode) & "')")
    lstCons.Sorted = False: lstCons.ListItems.Clear
    If Not (temprs.EOF And temprs.BOF) Then
        Listview_Loadval Me.lstCons.ListItems, temprs
        lstCons.Refresh
        lstCons.Enabled = True
    End If
End Sub

Sub FillCboSellingDealer()
    Dim rsSellingDealer                                As New ADODB.Recordset
    Set rsSellingDealer = gconDMIS.Execute("Select DealerName from CSMS_SellingDealer Order by DealerCode asc")
    If Not rsSellingDealer.EOF And Not rsSellingDealer.BOF Then
        Combo_Loadval cboSelling_Dealer, rsSellingDealer
    End If
    Set rsSellingDealer = Nothing
End Sub

Sub FillCboEndUser()
    Dim rsAllCustomer                                  As New ADODB.Recordset
    Set rsAllCustomer = gconDMIS.Execute("Select AcctName from All_Customer Where Custype = 'P' Order by AcctName asc")
    If Not rsAllCustomer.EOF And Not rsAllCustomer.BOF Then
        Combo_Loadval cboEndUser, rsAllCustomer
    End If
    Set rsAllCustomer = Nothing
End Sub

Private Sub cmd1_Click()
    picINS.Visible = True
End Sub

Private Sub cmdAdd_Click()
    'If Function_Access(LOGID, "Acess_ADD", "CUSTOMER VEHICLE") = False Then Exit Sub
    'UPDATED BY: ACL 010611
    'DESCRIPTION: USER ACCESS
    If Module_Access(LOGID, "CUSTOMER VEHICLE", "DATA ENTRY") = False Then Exit Sub
    AddorEdit = "ADD"
    cmdINS.Visible = False
    
    Call initMemvars
    picAdds.Visible = False
    picSaves.Visible = True
    Frame1.Enabled = True
    lstCons.Enabled = False
    Me.Caption = "Add Vehicle"
End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    picAdds.Visible = True
    picSaves.Visible = False
    Frame1.Enabled = False
    lstCons.Enabled = True
    Call StoreMemVars
    Me.Caption = "Customer Vehicle(s)"
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "CUSTOMER VEHICLE") = False Then Exit Sub
    On Error GoTo Errorcode:
    
    Dim rstmp                               As New ADODB.Recordset
    
    If MsgBox("Delete this Vehicle information" & vbCrLf & vbCrLf & "Are you sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    
    Set rstmp = gconDMIS.Execute("SELECT PLATE_NO FROM CSMS_REPOR WHERE PLATE_NO = " & N2Str2Null(txtPlateno) & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        MsgBox "you cannot delete this Vehicle Record. Vehicle information is already used in Billing system.", vbExclamation, "Info."
        Exit Sub
    End If
    Set rstmp = Nothing
    
    SQL_STATEMENT = "DELETE  from [CSMS_Cusveh] where id = " & Val(labID.Caption) & ""
    gconDMIS.Execute SQL_STATEMENT

    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("X", "CUSTOMER VEHICLE", SQL_STATEMENT, labID, "", "PLATE NO: " & txtPlateno, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call initMemvars
    Call FillSearchGrid
    Call StoreMemVars
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
        Dim rsLoad                                     As New ADODB.Recordset
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
    Me.Caption = "Edit Vehicle"
End Sub

Private Sub initMemvars()
    labVechicle.Caption = ""
    txtyear = ""
    txtMake = ""
    txtModel = ""
    txtEngine = ""
    txtPlateno = ""
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
    cboEndUser.Text = ""
    cboSelling_Dealer.Text = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'UPDATE BY: MJP 07182008 6:00 PM
Private Sub cmdINS_Click()
    Frame1.Enabled = False
    picAdds.Enabled = False
    picSaves.Enabled = False
    picINS.Visible = True
    picINS.ZOrder 0
    txtINSTYPE.SetFocus
End Sub
'UPDATE BY: MJP 07182008 6:00 PM

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "CUSTOMER VEHICLE") = False Then Exit Sub
    On Error GoTo Errorcode

    Screen.MousePointer = 11
    rptVehicle.ReportTitle = "Customer Vehicle Records"
    rptVehicle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptVehicle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    PrintSQLReport rptVehicle, CSMS_REPORT_PATH & "Customer Vehicle.rpt", "{All_Customer.Cuscde} = '" & labCustCode & "'", CSMS_REPORT_CONNECTION, 1

    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "CUSTOMER VEHICLE", "", labID, "", "PLATE NO: " & txtPlateno, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode:

    If txtMake.Text = "" Then
        ShowIsRequiredMsg "Make Cannot be Blank"
        On Error Resume Next
        txtMake.SetFocus
        Exit Sub
    End If

    If frmCSMSNewAppointment.labType(0).Caption = "Repair Order" Or frmCSMSNewAppointment.labType(0).Caption = "Estimate" Then
        If LTrim(RTrim(txtPlateno.Text)) = "" Then
            ShowIsRequiredMsg "Plate Number"
            On Error Resume Next
            txtPlateno.SetFocus
            Exit Sub
        End If

        If LTrim(RTrim(txtVIN.Text)) = "" Then
            ShowIsRequiredMsg "Vin no"
            On Error Resume Next
            txtVIN.SetFocus
            Exit Sub
        End If

        If LTrim(RTrim(labCustomer.Caption)) = "" Then
            ShowIsRequiredMsg "Customers Name"
            Exit Sub
        End If

        If LTrim(RTrim(txtColor.Text)) = "" Then
            If MsgBox("Color Information Missing. Do You Want to Continue", vbInformation + vbYesNo) = vbNo Then
                On Error Resume Next
                txtColor.SetFocus
                Exit Sub
            End If
        End If

        'If Not txtdateSold.Text = "" Then
        '    If IsDate(txtdateSold) = False Then
        '        MsgBox "Date Format is Incorrect", vbCritical, "ERROR"
        '        txtdateSold.SetFocus
        '        Exit Sub
        '    End If
        'End If

        'If Not txtDateDel.Text = "" Then
        '    If IsDate(txtDateDel) = False Then
        '        MsgBox "Date Format is Incorrect", vbCritical, "ERROR"
        '        txtDateDel.SetFocus
        '        Exit Sub
        '    End If
        'End If
    Else
        If LTrim(RTrim(txtPlateno.Text)) = "" Then
            ShowIsRequiredMsg "Plate Number"
            On Error Resume Next
            txtPlateno.SetFocus
            Exit Sub
        End If

        '        If LTrim(RTrim(txtVIN.Caption)) = "" Then
        '            ShowIsRequiredMsg "Vin no"
        '            On Error Resume Next
        '            txtVIN.SetFocus
        '            Exit Sub
        '        End If
    End If

    Dim xCuscde                                         As String
    Dim xNIYM                                           As String
    Dim XYEAR                                           As String
    Dim xMake                                           As String
    Dim xMODEL                                          As String
    Dim xENGINE                                         As String
    Dim xPLATE_NO                                       As String
    Dim xCLRCDE                                         As String
    Dim xSERIAL                                         As String
    Dim xPRODNO                                         As String
    Dim xTIN_NUMBER                                     As String
    Dim xWAR_CERT                                       As String
    Dim xKMreading                                      As String
    Dim xVIN                                            As String
    Dim xVCOND_NO                                       As String
    Dim xD_SOLD                                         As String
    Dim xDEL_DATE                                       As String
    Dim SQL                                             As String
    Dim EXIST                                           As Boolean
    Dim TMP_PLATE                                       As String
    Dim TMP_VIN                                         As String

    TMP_PLATE = txtPlateno.Text
    TMP_VIN = txtVIN.Text

    If AddorEdit = "ADD" Then
        If CheckIfPlateNoAlreadyExist("PLATE_NO", TMP_PLATE) Then
            MsgBox "Plate No. Already Exist", vbExclamation, "Customer Vehicle Info"
            txtPlateno.SetFocus
            Exit Sub
        End If
    Else
        If CheckIfPlateNoAlreadyExist("PLATE_NO", TMP_PLATE) Then
            If ReturnVehicleID(TMP_PLATE, "PLATE_NO", labID.Caption, "Plate no.") = True Then
                txtPlateno.SetFocus
                Exit Sub
            End If
        End If
    End If

    If Not txtVIN.Text = "" Then
        If AddorEdit = "ADD" Then
            If CheckIfPlateNoAlreadyExist("VIN", TMP_VIN) Then
                MsgBox "VIN No. Already Exist", vbExclamation, "Add Customer Vehicle"
                txtVIN.SetFocus
                Exit Sub
            End If
        Else
            If CheckIfPlateNoAlreadyExist("VIN", TMP_VIN) Then
                If ReturnVehicleID(TMP_VIN, "VIN", labID, "Vin no") = True Then
                    txtPlateno.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

    xCuscde = N2Str2Null(labCustCode)
    xNIYM = N2Str2Null(labCustomer.Caption)
    XYEAR = N2Str2Null(txtyear)
    xMake = N2Str2Null(txtMake)
    xMODEL = N2Str2Null(txtModel)
    xENGINE = N2Str2Null(txtEngine)
    xPLATE_NO = N2Str2Null(txtPlateno)
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

    If AddorEdit = "ADD" Then
        SQL = "Insert into CSMS_Cusveh "
        SQL = SQL & " (CUSCDE,NIYM,Yer,Make,MODEL,ENGINE,PLATE_NO,CLRCDE,SERIAL,PRODNO,TIN_NUMBER,WAR_CERT,KMReading,VIN,VCOND_NO,D_SOLD,Description,DEL_DATE,SELLING_DEALER,ENDUSER) VALUES("
        SQL = SQL & xCuscde & ","
        SQL = SQL & xNIYM & ","
        SQL = SQL & XYEAR & ","
        SQL = SQL & xMake & ","
        SQL = SQL & xMODEL & ","
        SQL = SQL & xENGINE & ","
        SQL = SQL & xPLATE_NO & ","
        SQL = SQL & xCLRCDE & ","
        SQL = SQL & xSERIAL & ","
        SQL = SQL & xPRODNO & ","
        SQL = SQL & xTIN_NUMBER & ","
        SQL = SQL & xWAR_CERT & ","
        SQL = SQL & xKMreading & ","
        SQL = SQL & xVIN & ","
        SQL = SQL & xVCOND_NO & ","
        SQL = SQL & xD_SOLD & ","
        SQL = SQL & N2Str2Null(txtDescription) & ","
        SQL = SQL & xDEL_DATE & ","
        SQL = SQL & N2Str2Null(SetSellingDealer(cboSelling_Dealer.Text, 2)) & ","
        SQL = SQL & N2Str2Null(SetEndUser(cboEndUser.Text, 2)) & ")"

        gconDMIS.Execute SQL
        SQL_STATEMENT = SQL

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("A", "CUSTOMER VEHICLE", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtPlateno), "PLATE_NO", "CSMS_CUSVEH"), "", "PLATE NO: " & txtPlateno, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        Call ShowSuccessFullyAdded
    Else
        'UPDATE BY : MJP 07182008 6:30 PM
        Dim getOldPlate                                As New ADODB.Recordset
        Dim rstmp                                      As New ADODB.Recordset
        Dim RSMAK                                      As New ADODB.Recordset
        Dim vREPOR                                     As String
        Dim vOldPlate                                  As String

        Set getOldPlate = gconDMIS.Execute("select plate_no from csms_cusveh where id = " & labID & "")
        If Not (getOldPlate.BOF And getOldPlate.EOF) Then
            vOldPlate = Null2String(getOldPlate!PLATE_NO)
        End If
        Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE PLATE_NO = '" & vOldPlate & "'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            If MsgBox("This vehicle had a past transaction. " & vbCrLf & "Editing this information will update all past records", vbQuestion + vbYesNo, "Are you sure") = vbNo Then
                Exit Sub
            Else
                SQL_STATEMENT = "update csms_repor set Plate_no = " & xPLATE_NO & ",MODEL = " & xMODEL & ",VIN = " & xVIN & " where plate_no = '" & vOldPlate & "'"
                gconDMIS.Execute SQL_STATEMENT

                Set RSMAK = gconDMIS.Execute("SELECT REP_OR,PLATE_NO FROM CSMS_REPOR WHERE PLATE_NO = '" & vOldPlate & "'")
                If Not (RSMAK.BOF And RSMAK.EOF) Then
                    Do While Not RSMAK.EOF
                        vREPOR = Null2String(RSMAK!REP_OR)
                        'NEW LOG AUDIT-----------------------------------------------------
                        Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREPOR), "REP_OR", "CSMS_REPOR"), "", "PLATE NO: " & vOldPlate, "", "")
                        'NEW LOG AUDIT-----------------------------------------------------
                        RSMAK.MoveNext
                    Loop
                End If
                Set RSMAK = Nothing

                gconDMIS.Execute "update csms_repairorder set Plate_no = " & xPLATE_NO & ",MODEL = " & xMODEL & " where plate_no = '" & vOldPlate & "'"
                gconDMIS.Execute "update csms_appointment set plate_no = " & xPLATE_NO & ",MODEL = " & xMODEL & " where plate_no = '" & vOldPlate & "'"
            End If
        End If
        'UPDATE BY : MJP 07182008 6:30 PM

        SQL = "UPDATE CSMS_Cusveh SET "
        SQL = SQL & " CUSCDE = " & xCuscde & ","
        SQL = SQL & " NIYM = " & xNIYM & ","
        SQL = SQL & " Yer = " & XYEAR & ","
        SQL = SQL & " Make = " & xMake & ","
        SQL = SQL & " MODEL = " & xMODEL & ","
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
        SQL = SQL & " Selling_Dealer = " & N2Str2Null(SetSellingDealer(cboSelling_Dealer, 2)) & ","
        SQL = SQL & " EndUser = " & N2Str2Null(SetEndUser(cboEndUser, 2)) & ","
        SQL = SQL & " DEL_DATE = " & xDEL_DATE
        SQL = SQL & " WHERE ID = " & labID
        gconDMIS.Execute SQL
        SQL_STATEMENT = SQL

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("E", "CUSTOMER VEHICLE", SQL_STATEMENT, labID, "", "PLATE NO: " & txtPlateno, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        Call ShowSuccessFullyUpdated
    End If

CONT:
    rsVehicle.Requery
    If AddorEdit = "EDIT" Then
        rsVehicle.Find ("ID = " & labID)
    End If

    Call FillSearchGrid
    cmdCancel.Value = True

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdSelect_Click()
    Set frmSelectMakeMode = New frmCSMSYrMkMlEgn
    frmSelectMakeMode.Show 1
End Sub

'UPDATE BY: MJP 07182008 6:00 PM
Private Sub Command1_Click()
    Frame1.Enabled = True
    picAdds.Enabled = True
    picSaves.Enabled = True
    picINS.Visible = False
    picINS.ZOrder 1
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
        SQL_STATEMENT = "UPDATE CSMS_CUSVEH SET INS_TYPE = " & vINSTYPE & ",INS_COMP = " & vINSCOMP & _
            ",INS_EXP_DATE = " & VINSDATE & ",FIN_TYPE = " & vFINTYPE & ",FIN_COMP = " & vFINCOMP & _
            ",FIN_EXP_DATE = " & vFINDATE & " WHERE ID = " & labID.Caption & ""
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("E", "CUSTOMER VEHICLE", SQL_STATEMENT, labID, "", "PLATE NO: " & txtPlateno, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        Call ShowSuccessFullyUpdated
    End If

    Call Command1_Click
End Sub
'UPDATE BY: MJP 07182008 6:00 PM

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call rsRefresh
    Call StoreMemVars
    Call FillCboSellingDealer
    Call FillCboEndUser
    Call FillSearchGrid
    Screen.MousePointer = 0
End Sub

Private Sub lstCons_DblClick()
    If lstCons.SelectedItem Is Nothing Then Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub lstCons_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsVehicle.MoveFirst
    rsVehicle.Find ("ID=" & ITEM.SubItems(15))
    Call StoreMemVars
End Sub

Private Sub txtConduction_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
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


Private Sub txtSerial_Change()
'txtVIN.Text = txtSerial.Text
 Call serial

End Sub

Sub serial()
txtVIN.Text = txtSerial.Text
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
