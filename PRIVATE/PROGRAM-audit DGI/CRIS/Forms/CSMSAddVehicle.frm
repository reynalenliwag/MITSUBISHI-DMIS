VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSAddVehicle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Vehicle"
   ClientHeight    =   8790
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
   Icon            =   "CSMSAddVehicle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox labCustomer 
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
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   74
      Top             =   510
      Width           =   7125
   End
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
      TabIndex        =   58
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
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   8985
      Begin VB.ComboBox cboSelling_Dealer 
         BackColor       =   &H00C0C000&
         Height          =   345
         Left            =   1020
         TabIndex        =   68
         Text            =   "Combo1"
         Top             =   3540
         Width           =   5535
      End
      Begin VB.ComboBox cboEndUser 
         BackColor       =   &H00C0C000&
         Height          =   345
         Left            =   1020
         TabIndex        =   67
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
         Height          =   555
         Left            =   150
         MaxLength       =   100
         TabIndex        =   69
         Top             =   4350
         Width           =   8685
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   2925
         Left            =   4920
         TabIndex        =   5
         Top             =   180
         Width           =   4035
         Begin VB.TextBox txtprdn 
            Height          =   330
            Left            =   1650
            MaxLength       =   6
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   60
            Width           =   2265
         End
         Begin VB.TextBox txtKMR 
            Height          =   330
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   1140
            Width           =   2265
         End
         Begin VB.TextBox txtVIN 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   20
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   1500
            Width           =   2265
         End
         Begin VB.TextBox txtTIN 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   15
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   420
            Width           =   2265
         End
         Begin VB.TextBox txtWCN 
            Height          =   330
            Left            =   1650
            MaxLength       =   15
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   780
            Width           =   2265
         End
         Begin VB.TextBox txtdateSold 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   2220
            Width           =   2265
         End
         Begin VB.TextBox txtDateDel 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   2580
            Width           =   2265
         End
         Begin VB.TextBox txtConduction 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1650
            MaxLength       =   8
            TabIndex        =   17
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
            TabIndex        =   6
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
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   9
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
            TabIndex        =   13
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
            TabIndex        =   19
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
            TabIndex        =   21
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
            TabIndex        =   16
            Top             =   1890
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2805
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   4725
         Begin VB.TextBox txtyear 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtMake 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   390
            Width           =   2475
         End
         Begin VB.TextBox txtModel 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   780
            Width           =   2475
         End
         Begin VB.TextBox txtEngine 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   900
            TabIndex        =   32
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
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   1560
            Width           =   2475
         End
         Begin VB.TextBox txtSerial 
            Height          =   330
            Left            =   900
            MaxLength       =   18
            TabIndex        =   39
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
            TabIndex        =   34
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
            TabIndex        =   37
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
            TabIndex        =   38
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   26
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
            TabIndex        =   30
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
            TabIndex        =   33
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
            TabIndex        =   35
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
            TabIndex        =   40
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
            TabIndex        =   41
            Top             =   1980
            Width           =   915
         End
         Begin VB.Label labID 
            Caption         =   "Label4"
            Height          =   285
            Left            =   3450
            TabIndex        =   29
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
            TabIndex        =   27
            Top             =   120
            Width           =   2205
         End
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   18
         Left            =   270
         TabIndex        =   72
         Top             =   3570
         Width           =   915
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   71
         Top             =   3180
         Width           =   915
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
         TabIndex        =   70
         Top             =   4050
         Width           =   2805
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   4440
      ScaleHeight     =   3315
      ScaleWidth      =   4155
      TabIndex        =   42
      Top             =   2550
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
         TabIndex        =   43
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
         TabIndex        =   45
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
         TabIndex        =   47
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
         TabIndex        =   49
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
         TabIndex        =   51
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
         TabIndex        =   53
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
         TabIndex        =   55
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
         TabIndex        =   57
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
         TabIndex        =   44
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
         TabIndex        =   46
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
         TabIndex        =   48
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
         TabIndex        =   50
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
         TabIndex        =   52
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
         TabIndex        =   54
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
         TabIndex        =   56
         Top             =   2400
         Width           =   1485
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7695
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   64
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
         MouseIcon       =   "CSMSAddVehicle.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   66
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
         MouseIcon       =   "CSMSAddVehicle.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Save New Vehicle"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   3120
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   59
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
         MouseIcon       =   "CSMSAddVehicle.frx":11FC
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":134E
         Style           =   1  'Graphical
         TabIndex        =   63
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
         MouseIcon       =   "CSMSAddVehicle.frx":16B4
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":1806
         Style           =   1  'Graphical
         TabIndex        =   73
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
         MouseIcon       =   "CSMSAddVehicle.frx":1B6C
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":1CBE
         Style           =   1  'Graphical
         TabIndex        =   62
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
         MouseIcon       =   "CSMSAddVehicle.frx":1FE9
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":213B
         Style           =   1  'Graphical
         TabIndex        =   61
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
         MouseIcon       =   "CSMSAddVehicle.frx":2497
         MousePointer    =   99  'Custom
         Picture         =   "CSMSAddVehicle.frx":25E9
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Add Vehicle"
         Top             =   60
         Width           =   705
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
      TabIndex        =   2
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
      Left            =   4170
      TabIndex        =   3
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
Dim AddorEdit                           As String
Dim rsVehicle                           As ADODB.Recordset
Public CustomerCode                     As String

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "CUSTOMER VEHICLE INFORMATION") = False Then Exit Sub
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
    StoreMemVars
    Me.Caption = "Customer Vehicle(s)"
End Sub

Private Sub cmdCustomer_Click()
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "CUSTOMER VEHICLE") = False Then Exit Sub
    On Error GoTo Errorcode:
    If MsgBox("Delete this item..." & vbCrLf & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, "Message Box") = vbNo Then
        Exit Sub
    End If
    gconDMIS.Execute "delete  from [CSMS_Cusveh] where id = " & Val(labid.Caption) & ""

    LogAudit "X", "CUSTOMER VEHICLE", "CUSTCODE " & labCustCode & "PLATE NO " & txtPlateno

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
    If Function_Access(LOGID, "Acess_Edit", "CUSTOMER VEHICLE INFORMATION") = False Then Exit Sub
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
    txtTin = ""
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
    cboEndUser.Text = ""
    cboSelling_Dealer.Text = ""
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

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "CUSTOMER VEHICLE INFORMATION") = False Then Exit Sub
    On Error GoTo Errorcode
    Screen.MousePointer = 11
    rptVehicle.ReportTitle = "Customer Vehicle Records"
    rptVehicle.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptVehicle.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptVehicle, CSMS_REPORT_PATH & "Customer Vehicle.rpt", "{All_Customer.Cuscde} = '" & labCustCode & "'", CSMS_REPORT_CONNECTION, 1
    LogAudit "V", "CUSTOMER VEHICLE RECORDS - REPORTS ", labCustCode.Caption
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()

    On Error GoTo Errorcode:

    If LTrim(RTrim(txtPlateno.Text)) = "" Then
        ShowIsRequiredMsg "Plate Number"
        On Error Resume Next
        txtPlateno.SetFocus
        Exit Sub
    End If

    If LTrim(RTrim(labCustomer)) = "" Then
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


    Dim xCuscde, xNIYM, xYear, xMake, xModel, xENGINE, xPLATE_NO, xCLRCDE, xSERIAL, xPRODNO, xTIN_NUMBER, xWAR_CERT, xKMreading, xVIN, xVCOND_NO, xD_SOLD, xDEL_DATE As String
    Dim SQL                             As String
    Dim EXIST                           As Boolean

    If AddorEdit = "ADD" Then
        Call CheckIfPlateNoAlreadyExist(EXIST)
        If EXIST Then
            MsgBox "Plate No. Already Exist", vbExclamation, "Add Customer Vehicle"
            txtPlateno.SetFocus
            Exit Sub
        End If
    End If

    xCuscde = N2Str2Null(labCustCode)
    xNIYM = N2Str2Null(labCustomer)
    xYear = N2Str2Null(txtyear)
    xMake = N2Str2Null(txtMake)
    xModel = N2Str2Null(txtModel)
    xENGINE = N2Str2Null(txtEngine)
    xPLATE_NO = N2Str2Null(txtPlateno)
    xCLRCDE = N2Str2Null(SetColor(txtColor))
    xSERIAL = N2Str2Null(txtSerial)
    xPRODNO = N2Str2Null(txtprdn)
    xTIN_NUMBER = N2Str2Null(txtTin)
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
        SQL = SQL & xYear & ","
        SQL = SQL & xMake & ","
        SQL = SQL & xModel & ","
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

        LogAudit "A", "CUSTOMER VEHICLE", "CUSTCODE" & xCuscde & "PLATE NO " & xPLATE_NO


    Else

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
        SQL = SQL & " Selling_Dealer=" & N2Str2Null(SetSellingDealer(cboSelling_Dealer, 2)) & ","
        SQL = SQL & " EndUser=" & N2Str2Null(SetEndUser(cboEndUser, 2)) & ","
        SQL = SQL & " DEL_DATE=" & xDEL_DATE
        SQL = SQL & " WHERE ID=" & labid
        gconDMIS.Execute SQL

        LogAudit "E", "CUSTOMER VEHICLE", "CUSTCODE" & xCuscde & "PLATE NO " & xPLATE_NO

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
    rsColor.Open "select COLOR_CODE,COLOR_DESC from ALL_Color where COLOR_DESC = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        SetColor = Null2String(rsColor!Color_Code)
    Else
        SetColor = ""
    End If
End Function

Function GetColor(CCC As String)
    Dim rsColor                         As ADODB.Recordset
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
        .labformname = "frmCSMSAddVehicle"
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
    FillCboSellingDealer
    FillCboEndUser
    FillSearchGrid
End Sub

Sub rsRefresh()
    Set rsVehicle = New ADODB.Recordset
    rsVehicle.Open "select * from CSMS_Cusveh where (CUSCDE = '" & CustomerCode & "' OR ENDUSER = '" & CustomerCode & "')", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not (rsVehicle.EOF And rsVehicle.BOF) Then
        txtyear = Null2String(rsVehicle!Yer)
        txtMake = Null2String(rsVehicle!Make)
        txtModel = Null2String(rsVehicle!Model)
        txtEngine = Null2String(rsVehicle!Engine)
        txtPlateno = Null2String(rsVehicle!Plate_no)
        txtColor = GetColor(Null2String(rsVehicle!ClrCde))
        txtSerial = Null2String(rsVehicle!Serial)
        txtprdn = Null2String(rsVehicle!ProdNo)
        txtTin = Null2String(rsVehicle!TIN_Number)
        txtWCN = Null2String(rsVehicle!War_Cert)
        txtKMR = NumericVal(rsVehicle!KMReading)
        txtVIN = Null2String(rsVehicle!Vin)
        txtConduction = Null2String(rsVehicle!VCond_No)
        txtdateSold = Null2String(rsVehicle!D_Sold)
        txtDateDel = Null2String(rsVehicle!Del_Date)
        txtDescription = Null2String(rsVehicle!Description)
        labid.Caption = rsVehicle!ID
        labVechicle.Caption = Trim(txtyear) & "  " & Trim(txtMake) & "  " & Trim(txtModel) & "  " & Trim(txtEngine)
        cboEndUser.Text = SetEndUser(Null2String(rsVehicle!ENDUSER), 1)
        cboSelling_Dealer.Text = SetSellingDealer(Null2String(rsVehicle!Selling_Dealer), 1)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub
Sub FillSearchGrid()
    Dim temprs                          As ADODB.Recordset

    lstCons.Enabled = False

    Set temprs = gconDMIS.Execute("select YER,MAKE,MODEL,ENGINE,PLATE_NO,CLRCDE,SERIAL,PRODNO,TIN_NUMBER,WAR_CERT,KMreading,VIN,VCOND_NO,D_SOLD,DEL_DATE,ID from CSMS_Cusveh where (CUSCDE = '" & Repleys(CustomerCode) & "' OR ENDUSER = '" & Repleys(CustomerCode) & "')")
    lstCons.Sorted = False: lstCons.ListItems.Clear
    If Not (temprs.EOF And temprs.BOF) Then
        Listview_Loadval Me.lstCons.ListItems, temprs
        lstCons.Refresh
        lstCons.Enabled = True
    End If
End Sub

Private Sub lstCons_DblClick()
    If lstCons.SelectedItem Is Nothing Then Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub lstCons_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsVehicle.MoveFirst
    rsVehicle.Find ("ID=" & Item.SubItems(15))
    StoreMemVars
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

Sub FillCboSellingDealer()
    Dim rsSellingDealer                 As ADODB.Recordset
    Set rsSellingDealer = New ADODB.Recordset
    Set rsSellingDealer = gconDMIS.Execute("Select DealerName from CSMS_SellingDealer Order by DealerCode asc")
    If Not rsSellingDealer.EOF And Not rsSellingDealer.BOF Then
        Combo_Loadval cboSelling_Dealer, rsSellingDealer
    End If
    Set rsSellingDealer = Nothing
End Sub

Function SetSellingDealer(XXX As String, CodeOrName As Integer) As String
    Dim rsSellingDealer                 As ADODB.Recordset
    Set rsSellingDealer = New ADODB.Recordset
    Dim SelectionCodeOrName             As String
    If CodeOrName = 1 Then
        SelectionCodeOrName = "DealerCode"
    Else
        SelectionCodeOrName = "DealerName"
    End If
    Set rsSellingDealer = gconDMIS.Execute("Select * from CSMS_SellingDealer Where " & SelectionCodeOrName & " = '" & XXX & "'")
    If Not rsSellingDealer.EOF And Not rsSellingDealer.BOF Then
        If CodeOrName = 1 Then
            SetSellingDealer = Null2String(rsSellingDealer!DealerName)
        Else
            SetSellingDealer = Null2String(rsSellingDealer!DealerCode)
        End If
    End If
End Function

Sub FillCboEndUser()
    Dim rsAllCustomer                   As ADODB.Recordset
    Set rsAllCustomer = New ADODB.Recordset
    Set rsAllCustomer = gconDMIS.Execute("Select AcctName from All_Customer Where Custype = 'P' Order by AcctName asc")
    If Not rsAllCustomer.EOF And Not rsAllCustomer.BOF Then
        Combo_Loadval cboEndUser, rsAllCustomer
    End If
    Set rsAllCustomer = Nothing
End Sub

Function SetEndUser(XXX As String, CodeOrName As Integer) As String
    Dim rsEndUser                       As ADODB.Recordset
    Set rsEndUser = New ADODB.Recordset
    Dim SelectionCodeOrName             As String
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
