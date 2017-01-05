VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSEditROVehicle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Vehicle"
   ClientHeight    =   6315
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
   Icon            =   "CSMSEditROVehicle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7665
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   75
      Top             =   5445
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
         MouseIcon       =   "CSMSEditROVehicle.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditROVehicle.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   77
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
         MouseIcon       =   "CSMSEditROVehicle.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditROVehicle.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   76
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
         MouseIcon       =   "CSMSEditROVehicle.frx":11FC
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditROVehicle.frx":134E
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
         MouseIcon       =   "CSMSEditROVehicle.frx":16B4
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditROVehicle.frx":1806
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
         MouseIcon       =   "CSMSEditROVehicle.frx":1B31
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditROVehicle.frx":1C83
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
         MouseIcon       =   "CSMSEditROVehicle.frx":1FDF
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditROVehicle.frx":2131
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
      TabIndex        =   38
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
         Left            =   60
         MaxLength       =   50
         TabIndex        =   18
         Top             =   3630
         Width           =   8775
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   2925
         Left            =   4920
         TabIndex        =   39
         Top             =   150
         Width           =   4035
         Begin VB.TextBox txtprdn 
            Height          =   330
            Left            =   1650
            MaxLength       =   6
            TabIndex        =   9
            Top             =   60
            Width           =   2265
         End
         Begin VB.TextBox txtKMR 
            Height          =   330
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   12
            Top             =   1140
            Width           =   2265
         End
         Begin VB.TextBox txtVIN 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   20
            TabIndex        =   13
            Top             =   1500
            Width           =   2265
         End
         Begin VB.TextBox txtTIN 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   15
            TabIndex        =   10
            Top             =   420
            Width           =   2265
         End
         Begin VB.TextBox txtWCN 
            Height          =   330
            Left            =   1650
            MaxLength       =   15
            TabIndex        =   11
            Top             =   780
            Width           =   2265
         End
         Begin VB.TextBox txtdateSold 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   15
            Top             =   2220
            Width           =   2265
         End
         Begin VB.TextBox txtDateDel 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   16
            Top             =   2580
            Width           =   2265
         End
         Begin VB.TextBox txtConduction 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   8
            TabIndex        =   14
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
            TabIndex        =   40
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
            TabIndex        =   42
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
            TabIndex        =   43
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
            TabIndex        =   41
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
            TabIndex        =   44
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
            TabIndex        =   46
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
            TabIndex        =   47
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
            TabIndex        =   45
            Top             =   1890
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2805
         Left            =   120
         TabIndex        =   48
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
            TabIndex        =   3
            Top             =   390
            Width           =   2475
         End
         Begin VB.TextBox txtModel 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   780
            Width           =   2475
         End
         Begin VB.TextBox txtEngine 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            TabIndex        =   5
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
            TabIndex        =   6
            Top             =   1560
            Width           =   2475
         End
         Begin VB.TextBox txtSerial 
            Height          =   330
            Left            =   900
            MaxLength       =   18
            TabIndex        =   8
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
            TabIndex        =   56
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
            TabIndex        =   7
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
            TabIndex        =   58
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   51
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
            TabIndex        =   54
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
            TabIndex        =   55
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
            TabIndex        =   57
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
            TabIndex        =   52
            Top             =   120
            Width           =   2205
         End
         Begin VB.Label labID 
            Caption         =   "Label4"
            Height          =   285
            Left            =   3480
            TabIndex        =   53
            Top             =   1260
            Visible         =   0   'False
            Width           =   1185
         End
      End
      Begin wizButton.cmd cmdINS 
         Height          =   435
         Left            =   5040
         TabIndex        =   17
         Top             =   3120
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   767
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
         MICON           =   "CSMSEditROVehicle.frx":2444
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
         TabIndex        =   78
         Top             =   3330
         Width           =   2805
      End
   End
   Begin VB.PictureBox picINS 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   2010
      ScaleHeight     =   3765
      ScaleWidth      =   5565
      TabIndex        =   79
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
         TabIndex        =   27
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
         TabIndex        =   28
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
         TabIndex        =   30
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
         TabIndex        =   31
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
         MouseIcon       =   "CSMSEditROVehicle.frx":2460
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditROVehicle.frx":25B2
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Cancel"
         Top             =   2880
         Width           =   705
      End
      Begin MSComCtl2.DTPicker dtpINS 
         Height          =   345
         Left            =   1920
         TabIndex        =   29
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
         Format          =   20643841
         CurrentDate     =   39647
      End
      Begin MSComCtl2.DTPicker dtpFIN 
         Height          =   345
         Left            =   1920
         TabIndex        =   32
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
         Format          =   20643841
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
         MouseIcon       =   "CSMSEditROVehicle.frx":28F0
         MousePointer    =   99  'Custom
         Picture         =   "CSMSEditROVehicle.frx":2A42
         Style           =   1  'Graphical
         TabIndex        =   33
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
      Height          =   3345
      Left            =   4890
      ScaleHeight     =   3315
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   21
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
         TabIndex        =   22
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
         TabIndex        =   23
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
         TabIndex        =   24
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
         TabIndex        =   25
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
         TabIndex        =   26
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
      TabIndex        =   35
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
      TabIndex        =   36
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
      TabIndex        =   37
      Top             =   720
      Width           =   5115
   End
End
Attribute VB_Name = "frmCSMSEditROVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADDOREDIT                                          As String
Dim rsVehicle                                          As ADODB.Recordset
Public CustomerCode                                    As String
Dim WithEvents frmSelectMakeMode                       As frmCSMSYrMkMlEgn
Attribute frmSelectMakeMode.VB_VarHelpID = -1


Private Sub frmSelectMakeMode_SelectedDetails(XYEAR As String, xMAKE As String, xMODEL As String, xENGINE As String, xModelDescription As String)
    txtyear = XYEAR
    txtMake = xMAKE
    txtModel = xMODEL
    txtEngine = xENGINE
    txtDescription = xModelDescription

End Sub



Function CheckIfPlateNoAlreadyExist() As Boolean
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select Plate_NO From CSMS_CusVeh Where Plate_NO = '" & txtPlateno.Text & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckIfPlateNoAlreadyExist = True
    Else
        CheckIfPlateNoAlreadyExist = False
    End If

    Set RSTMP = Nothing
End Function

Function ReturnVehicleID(vplate_no As String) As Integer
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("select ID from csms_Cusveh where plate_no = '" & vplate_no & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        ReturnVehicleID = RSTMP!ID
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

Sub StoreMemvars()
    If Not (rsVehicle.EOF And rsVehicle.BOF) Then
        txtyear = Null2String(rsVehicle!YER)
        txtMake = Null2String(rsVehicle!Make)
        txtModel = Null2String(rsVehicle!MODEL)
        txtEngine = Null2String(rsVehicle!Engine)
        txtPlateno = Null2String(rsVehicle!PLATE_NO)
        txtColor = GetColor(Null2String(rsVehicle!ClrCde))
        txtSerial = Null2String(rsVehicle!SERIAL)
        txtprdn = Null2String(rsVehicle!ProdNo)
        txtTIN = Null2String(rsVehicle!TIN_Number)
        txtWCN = Null2String(rsVehicle!War_Cert)
        txtKMR = NumericVal(rsVehicle!KMReading)
        txtVIN = Null2String(rsVehicle!Vin)
        txtConduction = Null2String(rsVehicle!VCond_no)
        txtdateSold = Null2String(rsVehicle!D_SOLD)
        txtDateDel = Null2String(rsVehicle!DEL_DATE)
        txtDescription = Null2String(rsVehicle!Description)
        labID.Caption = rsVehicle!ID
        labVechicle.Caption = Trim(txtyear) & "  " & Trim(txtMake) & "  " & Trim(txtModel) & "  " & Trim(txtEngine)
        txtINSTYPE.Text = Null2String(rsVehicle!INS_TYPE)
        txtINSCOMP.Text = Null2String(rsVehicle!INS_COMP)

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
    ADDOREDIT = "ADD"
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
    On Error GoTo ERRORCODE:
    If MsgBox("Delete this item..." & vbCrLf & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, "Message Box") = vbNo Then
        Exit Sub
    End If
    gconDMIS.Execute "delete  from [CSMS_Cusveh] where id = " & Val(labID.Caption) & ""
    StoreMemvars
    Exit Sub
ERRORCODE:
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
    ADDOREDIT = "EDIT"
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

Private Sub cmdSave_Click()
    On Error GoTo ERRORCODE:

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

    Dim xCUSCDE, xNIYM, XYEAR, xMAKE, xMODEL, xENGINE, xPLATE_NO, xCLRCDE, xSERIAL, xPRODNO, xTIN_NUMBER, xWAR_CERT, xKMreading, xVIN, xVCOND_NO, xD_SOLD, xDEL_DATE As String
    Dim SQL                                            As String

    Dim TMP_PLATE                                      As String

    TMP_PLATE = txtPlateno.Text

    If ADDOREDIT = "ADD" Then
        If CheckIfPlateNoAlreadyExist Then
            MsgBox "Plate No. Already Exist", vbExclamation, "Add Customer Vehicle"
            txtPlateno.SetFocus
            Exit Sub
        End If
    Else
        If CheckIfPlateNoAlreadyExist Then
            If Not labID.Caption = ReturnVehicleID(TMP_PLATE) Then
                MsgBox "Plate No. Already Exist", vbExclamation, "Add Customer Vehicle"
                txtPlateno.SetFocus
                Exit Sub
            End If
        End If
    End If

    xCUSCDE = N2Str2Null(labCustCode)
    xNIYM = N2Str2Null(labCustomer.Caption)
    XYEAR = N2Str2Null(txtyear)
    xMAKE = N2Str2Null(txtMake)
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

    Dim getOldPlate                                    As New ADODB.Recordset
    Dim RSTMP                                          As New ADODB.Recordset
    Dim vOldPlate                                      As String


    Set getOldPlate = gconDMIS.Execute("select plate_no from csms_cusveh where id = " & labID & "")
    If Not (getOldPlate.BOF And getOldPlate.EOF) Then
        vOldPlate = Null2String(getOldPlate!PLATE_NO)
    End If
    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE PLATE_NO = '" & vOldPlate & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If MsgBox("This vehicle had a past transaction. " & vbCrLf & "editing this information will upadte all past records", vbQuestion + vbYesNo, "Are you sure") = vbNo Then
            GoTo CONT
        Else
            gconDMIS.Execute "update csms_repor set Plate_no =" & xPLATE_NO & " where plate_no ='" & vOldPlate & "'"
            gconDMIS.Execute "update csms_repairorder set Plate_no =" & xPLATE_NO & " where plate_no ='" & vOldPlate & "'"
            gconDMIS.Execute "update csms_appointment set plate_no = " & xPLATE_NO & " where plate_no = '" & vOldPlate & "'"
        End If
    End If

    SQL = "UPDATE CSMS_Cusveh SET "
    SQL = SQL & " CUSCDE = " & xCUSCDE & ","
    SQL = SQL & " NIYM = " & xNIYM & ","
    SQL = SQL & " Yer = " & XYEAR & ","
    SQL = SQL & " Make = " & xMAKE & ","
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
    SQL = SQL & " DEL_DATE = " & xDEL_DATE
    SQL = SQL & " WHERE ID = " & labID
    gconDMIS.Execute SQL
    frmCSMSEditRO.SetVehicleInfo txtPlateno.Text

    LogAudit "E", "EDIT CUSTOMER VEHICLE INFORMATION", "CODE/PLATE " & labCustCode & "/" & txtPlateno

CONT:
    cmdCancel.Value = True
    Exit Sub

ERRORCODE:
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
                          ",FIN_EXP_DATE = " & vFINDATE & " WHERE ID = " & labID.Caption & "")

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
    StoreMemvars
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

