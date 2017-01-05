VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllCustomerVehicles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Vehicle"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
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
   Icon            =   "CustomerVehicles.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   3000
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   58
      Top             =   7110
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
         MouseIcon       =   "CustomerVehicles.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CustomerVehicles.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   62
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
         MouseIcon       =   "CustomerVehicles.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "CustomerVehicles.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   61
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
         MouseIcon       =   "CustomerVehicles.frx":11FF
         MousePointer    =   99  'Custom
         Picture         =   "CustomerVehicles.frx":1351
         Style           =   1  'Graphical
         TabIndex        =   60
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
         MouseIcon       =   "CustomerVehicles.frx":16AD
         MousePointer    =   99  'Custom
         Picture         =   "CustomerVehicles.frx":17FF
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Add Vehicle"
         Top             =   45
         Width           =   705
      End
   End
   Begin MSComctlLib.ListView lstCons 
      Height          =   2625
      Left            =   120
      TabIndex        =   57
      Top             =   4470
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   4630
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
      Height          =   3315
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   8985
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   3165
         Left            =   4920
         TabIndex        =   6
         Top             =   150
         Width           =   4035
         Begin VB.TextBox txtprdn 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1650
            MaxLength       =   30
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   30
            Width           =   2265
         End
         Begin VB.TextBox txtKMR 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1650
            MaxLength       =   30
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   1185
            Width           =   2265
         End
         Begin VB.TextBox txtVIN 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1650
            MaxLength       =   30
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   1570
            Width           =   2265
         End
         Begin VB.TextBox txtTIN 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1650
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   420
            Width           =   2265
         End
         Begin VB.TextBox txtWCN 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1650
            MaxLength       =   30
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   800
            Width           =   2265
         End
         Begin VB.TextBox txtdateSold 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1650
            MaxLength       =   30
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   2340
            Width           =   2265
         End
         Begin VB.TextBox txtDateDel 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1650
            MaxLength       =   30
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   2700
            Width           =   2265
         End
         Begin VB.TextBox txtConduction 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1650
            MaxLength       =   30
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   1955
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
            Top             =   90
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
            Top             =   1200
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
            Top             =   1590
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
            Top             =   2370
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
            Top             =   2730
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
            Top             =   1950
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
         Begin VB.ComboBox cboColor 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   900
            TabIndex        =   66
            Text            =   "Combo1"
            Top             =   1950
            Width           =   2475
         End
         Begin VB.TextBox txtyear 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtMake 
            BackColor       =   &H8000000F&
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
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   390
            Width           =   2475
         End
         Begin VB.TextBox txtModel 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   780
            Width           =   2475
         End
         Begin VB.TextBox txtEngine 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   33
            Text            =   "Text1"
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
            Height          =   360
            Left            =   900
            MaxLength       =   6
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   1560
            Width           =   2475
         End
         Begin VB.TextBox txtSerial 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   900
            MaxLength       =   18
            TabIndex        =   38
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
            TabIndex        =   39
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
            TabIndex        =   40
            Top             =   1980
            Width           =   915
         End
         Begin VB.Label labID 
            Caption         =   "Label4"
            Height          =   285
            Left            =   3450
            TabIndex        =   30
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
            TabIndex        =   28
            Top             =   120
            Width           =   2205
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   4440
      ScaleHeight     =   3315
      ScaleWidth      =   4155
      TabIndex        =   41
      Top             =   2790
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
         TabIndex        =   42
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
         TabIndex        =   44
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
         TabIndex        =   46
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
         TabIndex        =   48
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
         TabIndex        =   50
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
         TabIndex        =   52
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
         TabIndex        =   54
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
         TabIndex        =   56
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
         TabIndex        =   43
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
         TabIndex        =   45
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
         TabIndex        =   47
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
         TabIndex        =   49
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
         TabIndex        =   51
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
         TabIndex        =   53
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
         TabIndex        =   55
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
      TabIndex        =   63
      Top             =   7125
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
         MouseIcon       =   "CustomerVehicles.frx":1B12
         MousePointer    =   99  'Custom
         Picture         =   "CustomerVehicles.frx":1C64
         Style           =   1  'Graphical
         TabIndex        =   65
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
         MouseIcon       =   "CustomerVehicles.frx":1FA2
         MousePointer    =   99  'Custom
         Picture         =   "CustomerVehicles.frx":20F4
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Save New Vehicle"
         Top             =   30
         Width           =   705
      End
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
      Top             =   960
      Width           =   5115
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1365
   End
   Begin VB.Label labCustomer 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1890
      TabIndex        =   2
      Top             =   540
      Width           =   7215
   End
   Begin VB.Label labCustCode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1890
      TabIndex        =   1
      Top             =   60
      Width           =   1305
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
      Left            =   330
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmAllCustomerVehicles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                           As String
Dim rsVehicle                           As ADODB.Recordset
Public CustomerCode                     As String

Private Sub cmdAdd_Click()

    AddorEdit = "ADD"
    initMemvars
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
    StoreMemvars
    Me.Caption = "Customer Vehicle(s)"
End Sub

Private Sub cmdCustomer_Click()
    '  frmALLCustomer.Show 1
End Sub
Private Sub cmdDelete_Click()

    On Error GoTo Errorcode:
    If MsgBox("Delete this item..." & vbCrLf & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, "Message Box") = vbNo Then
        Exit Sub
    End If
    gconDMIS.Execute "delete  from [CSMS_Cusveh] where id = " & Val(labid.Caption) & ""
    StoreMemvars
    LogAudit "X", "CUSTOMER VEHICLE INFORMATION", "CUSTOMER CODE:" & labCustCode & " MODEL " & txtModel & " PLATE NO " & txtPlateno
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
        Dim rsLoad                      As ADODB.Recordset
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

    AddorEdit = "EDIT"
    picAdds.Visible = False
    picSaves.Visible = True
    Frame1.Enabled = True
    lstCons.Enabled = False

End Sub


Private Sub initMemvars()
    labVechicle.Caption = ""
    txtYear = ""
    txtMake = ""
    txtModel = ""
    txtEngine = ""
    txtPlateno = ""
    cboColor = ""
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
    Frame1.Enabled = False
    lstCons.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Sub CheckIfPlateNoAlreadyExist(EXIST As Boolean)
    Dim rsTmp                           As New ADODB.Recordset

    Set rsTmp = gconDMIS.Execute("Select Plate_NO From CSMS_CusVeh Where Plate_NO = '" & txtPlateno.Text & "'")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        EXIST = True
    Else
        EXIST = False
    End If

    Set rsTmp = Nothing
End Sub

Private Sub cmdSave_Click()

    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    If LTrim(RTrim(txtPlateno.Text)) = "" Then
        ShowIsRequiredMsg "Plate Number"
        On Error Resume Next
        txtPlateno.SetFocus
        Exit Sub
    End If

    If LTrim(RTrim(labCustomer.Caption)) = "" Then
        ShowIsRequiredMsg "Customers Name"
        Exit Sub
    End If

    Dim xCuscde, xNIYM, xYear, xMake, xModel, xENGINE, xPLATE_NO, xCLRCDE, xSERIAL, xPRODNO, xTIN_NUMBER, xWAR_CERT, xKMreading, xVIN, xVCOND_NO, xD_SOLD, xDEL_DATE As String
    Dim sql                             As String
    Dim lng                             As Integer
    '''''''AXP5/11/200715:02
    lng = gconDMIS.Execute("Select count(Plate_NO) From CSMS_CusVeh Where Plate_NO=" & N2Str2Null(txtPlateno)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MsgBox "Plate No. Already Exist", vbExclamation, "Add Customer Vehicle"
            txtPlateno.SetFocus
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsVehicle!Plate_no)) <> UCase(txtPlateno) Then
            MsgBox "Plate No. Already Exist", vbExclamation, "Add Customer Vehicle"
            txtPlateno.SetFocus
            Exit Sub
        End If
    End If




    xCuscde = N2Str2Null(CustomerCode)
    xNIYM = N2Str2Null(labCustomer.Caption)
    xYear = N2Str2Null(txtYear)
    xMake = N2Str2Null(txtMake)
    xModel = N2Str2Null(txtModel)
    xENGINE = N2Str2Null(txtEngine)
    xPLATE_NO = N2Str2Null(txtPlateno)
    xCLRCDE = N2Str2Null(SetColor(cboColor))
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
        sql = "Insert into CSMS_Cusveh "
        sql = sql & " (CUSCDE,NIYM,Yer,Make,MODEL,ENGINE,PLATE_NO,CLRCDE,SERIAL,PRODNO,TIN_NUMBER,WAR_CERT,KMReading,VIN,VCOND_NO,D_SOLD,DEL_DATE) VALUES("
        sql = sql & xCuscde & ","
        sql = sql & xNIYM & ","
        sql = sql & xYear & ","
        sql = sql & xMake & ","
        sql = sql & xModel & ","
        sql = sql & xENGINE & ","
        sql = sql & xPLATE_NO & ","
        sql = sql & xCLRCDE & ","
        sql = sql & xSERIAL & ","
        sql = sql & xPRODNO & ","
        sql = sql & xTIN_NUMBER & ","
        sql = sql & xWAR_CERT & ","
        sql = sql & xKMreading & ","
        sql = sql & xVIN & ","
        sql = sql & xVCOND_NO & ","
        sql = sql & xD_SOLD & ","
        sql = sql & xDEL_DATE & ")"

        gconDMIS.Execute sql
        LogAudit "A", "CUSTOMER VEHICLE INFORMATION", "CUSTOMER CODE:" & labCustCode & " MODEL " & txtModel & " PLATE NO " & txtPlateno
    Else

        sql = "UPDATE CSMS_Cusveh SET "
        sql = sql & " CUSCDE=" & xCuscde & ","
        sql = sql & " NIYM=" & xNIYM & ","
        sql = sql & " Yer=" & xYear & ","
        sql = sql & " Make=" & xMake & ","
        sql = sql & " MODEL=" & xModel & ","
        sql = sql & " ENGINE=" & xENGINE & ","
        sql = sql & " PLATE_NO=" & xPLATE_NO & ","
        sql = sql & " CLRCDE=" & xCLRCDE & ","
        sql = sql & " SERIAL=" & xSERIAL & ","
        sql = sql & " PRODNO=" & xPRODNO & ","
        sql = sql & " TIN_NUMBER=" & xTIN_NUMBER & ","
        sql = sql & " WAR_CERT=" & xWAR_CERT & ","
        sql = sql & " KMReading=" & xKMreading & ","
        sql = sql & " VIN=" & xVIN & ","
        sql = sql & " VCOND_NO=" & xVCOND_NO & ","
        sql = sql & " D_SOLD=" & xD_SOLD & ","
        sql = sql & " DEL_DATE=" & xDEL_DATE
        sql = sql & " WHERE ID=" & labid
        gconDMIS.Execute sql
        LogAudit "E", "CUSTOMER VEHICLE INFORMATION", "CUSTOMER CODE:" & labCustCode & " MODEL " & txtModel & " PLATE NO " & txtPlateno
    End If

    rsVehicle.Requery
    If AddorEdit = "EDIT" Then
        rsVehicle.Find ("ID=" & labid)
    End If

    Call FillSearchGrid
    cmdCancel.Value = True

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Function SetColor(CCC As String)
    Dim rsColor                         As ADODB.Recordset
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select COLOR_CODE from ALL_Color where COLOR_DESC = " & N2Str2Null(CCC), gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        SetColor = Null2String(rsColor!Color_code)
    Else
        SetColor = ""
    End If
End Function

Function GetColor(CCC As String)
    Dim rsColor                         As ADODB.Recordset
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select COLOR_DESC from ALL_Color where COLOR_CODE = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        GetColor = Null2String(rsColor!COLOR_DESC)
    Else
        GetColor = ""
    End If
End Function




Private Sub cmdSelect_Click()
    frmCSMSYrMkMlEgn.Show 1
End Sub


Private Sub Command3_Click()
    Picture3.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    initMemvars
    rsRefresh

    FillSearchGrid
    Combo_Loadval cboColor, gconDMIS.Execute("Select UPPER(COLOR_DESC) from ALL_COLOR")
    StoreMemvars
End Sub

Sub rsRefresh()
    Set rsVehicle = New ADODB.Recordset
    Call rsVehicle.Open("select YER,MAKE,MODEL,ENGINE,PLATE_NO,CLRCDE,SERIAL,PRODNO,TIN_NUMBER,WAR_CERT,KMreading,VIN,VCOND_NO,D_SOLD,DEL_DATE,ID from CSMS_Cusveh where CUSCDE = '" & CustomerCode & "'", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub


Sub StoreMemvars()

    If Not (rsVehicle.EOF And rsVehicle.BOF) Then
        txtYear = Null2String(rsVehicle!yER)
        txtMake = Null2String(rsVehicle!Make)
        txtModel = Null2String(rsVehicle!Model)
        txtEngine = Null2String(rsVehicle!Engine)
        txtPlateno = Null2String(rsVehicle!Plate_no)
        cboColor = GetColor(Null2String(rsVehicle!CLRCDE))
        txtSerial = Null2String(rsVehicle!Serial)
        txtprdn = Null2String(rsVehicle!ProdNo)
        txtTIN = Null2String(rsVehicle!tin_number)
        txtWCN = Null2String(rsVehicle!WAR_CERT)
        txtKMR = NumericVal(rsVehicle!KMReading)
        txtVIN = Null2String(rsVehicle!VIN)
        txtConduction = Null2String(rsVehicle!vcond_no)
        txtdateSold = Null2String(rsVehicle!D_Sold)
        txtDateDel = Null2String(rsVehicle!del_date)

        labid.Caption = rsVehicle!ID
        labVechicle.Caption = Trim(txtYear) & "  " & Trim(txtMake) & "  " & Trim(txtModel) & "  " & Trim(txtEngine)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub
Sub FillSearchGrid()
    Dim TempRs                          As ADODB.Recordset

    lstCons.Enabled = False

    Set TempRs = gconDMIS.Execute("select YER,MAKE,MODEL,ENGINE,PLATE_NO,CLRCDE,SERIAL,PRODNO,TIN_NUMBER,WAR_CERT,KMreading,VIN,VCOND_NO,D_SOLD,DEL_DATE,ID from CSMS_Cusveh where CUSCDE = '" & Repleys(CustomerCode) & "'")
    lstCons.Sorted = False: lstCons.ListItems.Clear
    If Not (TempRs.EOF And TempRs.BOF) Then
        Listview_Loadval Me.lstCons.ListItems, TempRs
        lstCons.Refresh
        lstCons.Enabled = True
    End If



End Sub

Private Sub lstCons_DblClick()
    If lstCons.SelectedItem Is Nothing Then Exit Sub
    cmdEdit.Value = True
    Me.Caption = "Edit Vehicle"
End Sub

Private Sub lstCons_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsVehicle.MoveFirst
    rsVehicle.Find ("ID=" & Item.SubItems(15))
    StoreMemvars
    '    txtyear = Item.Text
    '    txtMake = Item.SubItems(1)
    '    txtModel = Item.SubItems(2)
    '    txtEngine = Item.SubItems(3)
    '    txtPlateno = Item.SubItems(4)
    '    cboColor = GetColor(Item.SubItems(5))
    '    txtSerial = Item.SubItems(6)
    '    txtprdn = Item.SubItems(7)
    '    txtTIN = Item.SubItems(8)
    '    txtWCN = Item.SubItems(9)
    '    txtKMR = Item.SubItems(10)
    '    txtVIN = Item.SubItems(11)
    '    txtConduction = Item.SubItems(12)
    '    txtdateSold = Item.SubItems(13)
    '    txtDateDel = Item.SubItems(14)
    '    labID.Caption = Item.SubItems(15)
    '    labVechicle.Caption = Trim(txtyear) & "  " & Trim(txtMake) & "  " & Trim(txtModel) & "  " & Trim(txtEngine)
    '    il = Item.Index
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
