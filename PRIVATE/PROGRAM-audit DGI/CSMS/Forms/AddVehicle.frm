VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMSAddVehicle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Vehicle"
   ClientHeight    =   8220
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
   Icon            =   "AddVehicle.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   4650
      ScaleHeight     =   3315
      ScaleWidth      =   4155
      TabIndex        =   41
      Top             =   90
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
         Top             =   2790
         Width           =   1395
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         Height          =   3315
         Left            =   0
         Top             =   0
         Width           =   4155
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   2400
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdCustomer 
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
      Height          =   405
      Left            =   8760
      TabIndex        =   33
      Top             =   510
      Width           =   375
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
      Height          =   3225
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   8985
      Begin VB.CommandButton cmdSelect 
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
         Left            =   2220
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   330
         Width           =   465
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
         Left            =   3690
         TabIndex        =   40
         Top             =   2340
         Width           =   345
      End
      Begin VB.TextBox txtConduction 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6600
         MaxLength       =   8
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   2040
         Width           =   2265
      End
      Begin VB.TextBox txtColor 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   35
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
         Left            =   3690
         TabIndex        =   27
         Top             =   1560
         Width           =   765
      End
      Begin VB.TextBox txtDateDel 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   2760
         Width           =   2265
      End
      Begin VB.TextBox txtdateSold 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   2400
         Width           =   2265
      End
      Begin VB.TextBox txtWCN 
         Height          =   330
         Left            =   6600
         MaxLength       =   15
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   960
         Width           =   2265
      End
      Begin VB.TextBox txtTIN 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6600
         MaxLength       =   15
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   600
         Width           =   2265
      End
      Begin VB.TextBox txtVIN 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6600
         MaxLength       =   20
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1680
         Width           =   2265
      End
      Begin VB.TextBox txtKMR 
         Height          =   330
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1320
         Width           =   2265
      End
      Begin VB.TextBox txtprdn 
         Height          =   330
         Left            =   6600
         MaxLength       =   6
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   240
         Width           =   2265
      End
      Begin VB.TextBox txtSerial 
         Height          =   330
         Left            =   1170
         MaxLength       =   18
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2730
         Width           =   2475
      End
      Begin VB.TextBox txtPlateno 
         Height          =   330
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1950
         Width           =   2475
      End
      Begin VB.TextBox txtEngine 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1560
         Width           =   2475
      End
      Begin VB.TextBox txtModel 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1170
         Width           =   2475
      End
      Begin VB.TextBox txtMake 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   780
         Width           =   2475
      End
      Begin VB.TextBox txtyear 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   390
         Width           =   975
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
         Left            =   2790
         TabIndex        =   63
         Top             =   510
         Width           =   2145
      End
      Begin VB.Label labID 
         Caption         =   "Label4"
         Height          =   285
         Left            =   2250
         TabIndex        =   39
         Top             =   780
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Conduction No."
         Height          =   255
         Index           =   14
         Left            =   5220
         TabIndex        =   38
         Top             =   2070
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Color"
         Height          =   255
         Index           =   13
         Left            =   630
         TabIndex        =   36
         Top             =   2370
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Date Delivered"
         Height          =   255
         Index           =   12
         Left            =   5280
         TabIndex        =   25
         Top             =   2790
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Date Sold"
         Height          =   255
         Index           =   11
         Left            =   5640
         TabIndex        =   23
         Top             =   2430
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Warranty Cert. No."
         Height          =   255
         Index           =   10
         Left            =   5010
         TabIndex        =   21
         Top             =   990
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "TIN No."
         Height          =   255
         Index           =   9
         Left            =   5880
         TabIndex        =   19
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "VIN No."
         Height          =   255
         Index           =   8
         Left            =   5880
         TabIndex        =   17
         Top             =   1710
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Kilometer Reading"
         Height          =   255
         Index           =   7
         Left            =   4950
         TabIndex        =   15
         Top             =   1350
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Production No."
         Height          =   255
         Index           =   6
         Left            =   5280
         TabIndex        =   7
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Serial No."
         Height          =   255
         Index           =   5
         Left            =   270
         TabIndex        =   6
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Plate No."
         Height          =   255
         Index           =   4
         Left            =   330
         TabIndex        =   5
         Top             =   1980
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Engine"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   4
         Top             =   1590
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Model"
         Height          =   255
         Index           =   2
         Left            =   570
         TabIndex        =   3
         Top             =   1230
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Make"
         Height          =   255
         Index           =   1
         Left            =   630
         TabIndex        =   2
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Year"
         Height          =   255
         Index           =   0
         Left            =   690
         TabIndex        =   1
         Top             =   450
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lstCons 
      Height          =   2625
      Left            =   120
      TabIndex        =   34
      Top             =   4620
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
   Begin VB.CommandButton cmdQuit 
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
      Left            =   8400
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   7320
      Width           =   735
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
      Left            =   7680
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   7320
      Width           =   735
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
      Left            =   6960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   7320
      Width           =   735
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
      Left            =   6240
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   7320
      Width           =   735
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
      Left            =   5520
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   7320
      Width           =   735
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
      TabIndex        =   32
      Top             =   960
      Width           =   5115
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Code"
      Height          =   285
      Left            =   360
      TabIndex        =   31
      Top             =   90
      Width           =   1365
   End
   Begin VB.Label labCustomer 
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   30
      Top             =   510
      Width           =   6825
   End
   Begin VB.Label labCustCode 
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   29
      Top             =   90
      Width           =   1305
   End
   Begin VB.Label Label3 
      Caption         =   "Customer Name"
      Height          =   285
      Left            =   360
      TabIndex        =   28
      Top             =   540
      Width           =   1575
   End
End
Attribute VB_Name = "frmCSMSAddVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()


    cmdSave.Enabled = True: cmdAdd.Enabled = False
    cmdSelect.Enabled = True: Frame1.Enabled = True
End Sub

Private Sub cmdCustomer_Click()
    frmAllCustomer.Show 1
End Sub


Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete") = False Then Exit Sub

    If MsgBox("Delete this item..." & vbCrLf & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, "Message Box") = vbNo Then
        Exit Sub
    End If
    gconDMIS.Execute "delete  from [CSMS_Cusveh] where id = " & Val(labID.Caption) & ""
    LogAudit "X", "CUSTOMERVEHICLES", txtPlateno
    ShowVehicle
End Sub

Private Sub cmdDetails_Click()
    Picture3.Visible = True
    Dim rsLoad                         As ADODB.Recordset
    Set rsLoad = New ADODB.Recordset
    Set rsLoad = gconDMIS.Execute("Select * from All_Engine where ENGINE = '" & Trim(txtEngine) & "'")
    If Not rsLoad.EOF And Not rsLoad.BOF Then
        txtEnginetype = Null2String(rsLoad![engine])
        txtLiters = Null2String(rsLoad![Liters])
        txtCubic = Null2String(rsLoad![Cubic])
        txtDisplacement = Null2String(rsLoad![Displacement])
        txtFuelType = Null2String(rsLoad![FuelType])
        cboAspiration.Text = Null2String(rsLoad![Aspiration])
        txtEngineVIN = Null2String(rsLoad![EngineVIN])
    End If
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit") = False Then Exit Sub

    cmdEdit.Enabled = False: cmdSave.Enabled = True: cmdDelete.Enabled = False
    Frame1.Enabled = True
    lstCons.Enabled = False
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    labVechicle.Caption = ""
    txtYear = "": txtMake = "": txtModel = "": txtEngine = "": txtPlateno = "": txtColor = ""
    txtSerial = "": txtprdn = "": txtTIN = "": txtWCN = "": txtKMR = "": txtVIN = "": txtConduction = "": labID.Caption = ""
    txtdateSold = "": txtDateDel = "": cmdSelect.Enabled = False: Frame1.Enabled = False
    cmdAdd.Enabled = True: cmdEdit.Enabled = False: cmdDelete.Enabled = False: cmdSave.Enabled = False
    lstCons.Enabled = True
End Sub

Private Sub cmdSave_Click()
    If IsDate(txtdateSold) = False Then
        txtdateSold = ""
    End If
    If IsDate(txtDateDel) = False Then
        txtDateDel = ""
    End If
    If txtPlateno = "" Then
        MsgBox "No plate no."
        Exit Sub
    End If
    If txtPlateno = "" Then
        MsgBox "No plate no."
        Exit Sub
    End If
    If labCustomer.Caption = "" Then
        MsgBox "No customer name"
        Exit Sub
    End If
    Dim xCUSCDE, xNIYM, xYear, xMake, xModel, xENGINE, xPLATE_NO, xCLRCDE, xSERIAL, xPRODNO, xTIN_NUMBER, xWAR_CERT, xKMreading, xVIN, xVCOND_NO, xD_SOLD, xDEL_DATE As String
    xCUSCDE = N2Str2Null(labCustCode.Caption)
    xNIYM = N2Str2Null(labCustomer.Caption)
    xYear = N2Str2Null(txtYear)
    xMake = N2Str2Null(txtMake)
    xModel = N2Str2Null(txtModel)
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
    gconDMIS.Execute "delete from [CSMS_Cusveh] where id = " & Val(labID.Caption) & ""
    gconDMIS.Execute "Insert into CSMS_Cusveh " & _
                   " (CUSCDE,NIYM,Yer,Make,MODEL,ENGINE,PLATE_NO,CLRCDE,SERIAL,PRODNO,TIN_NUMBER,WAR_CERT,KMReading,VIN,VCOND_NO,D_SOLD,DEL_DATE)" & _
                   " values(" & xCUSCDE & "," & xNIYM & "," & xYear & "," & xMake & "," & xModel & "," & xENGINE & "," & xPLATE_NO & "," & xCLRCDE & "," & xSERIAL & "," & xPRODNO & "," & xTIN_NUMBER & "," & xWAR_CERT & "," & xKMreading & "," & xVIN & "," & xVCOND_NO & "," & xD_SOLD & "," & xDEL_DATE & ")"
    cmdSave.Enabled = False
    LogAudit "A", "CUSTOMERVEHICLES"
    ShowVehicle
End Sub

Function SetColor(CCC As String)
    Dim rsColor                        As ADODB.Recordset
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select COLOR_CODE,COLOR_DESC from ALL_Color where COLOR_DESC = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        SetColor = Null2String(rsColor!COLOR_CODE)
    Else
        SetColor = ""
    End If
End Function

Private Sub cmdSelect_Click()
    frmCSMSYrMkMlEgn.Show 1
End Sub

Private Sub Command2_Click()
    frmCSMSGetColor.Show 1
End Sub

Private Sub Command3_Click()
    Picture3.Visible = False
End Sub

Private Sub Form_Activate()
    ShowVehicle
End Sub

Private Sub Form_Load()
    cmdRefresh_Click
    labCustCode.Caption = "": labCustomer.Caption = ""
End Sub
Sub ShowVehicle()
    Dim rsVehicle                      As ADODB.Recordset
    Set rsVehicle = New ADODB.Recordset
    Set rsVehicle = gconDMIS.Execute("select YER,MAKE,MODEL,ENGINE,PLATE_NO,CLRCDE,SERIAL,PRODNO,TIN_NUMBER,WAR_CERT,KMreading,VIN,VCOND_NO,D_SOLD,DEL_DATE,ID from CSMS_Cusveh where CUSCDE = '" & labCustCode & "'")
    lstCons.Sorted = False: lstCons.ListItems.Clear
    If Not (rsVehicle.EOF And rsVehicle.BOF) Then
        Listview_Loadval Me.lstCons.ListItems, rsVehicle
        lstCons.Enabled = True
        lstCons.Refresh
    Else
        lstCons.Enabled = False
    End If
End Sub
Private Sub lstCons_DblClick()
    cmdAdd.Enabled = False: cmdSave.Enabled = False
    cmdEdit.Enabled = True: cmdDelete.Enabled = True
    txtYear = lstCons.SelectedItem
    txtMake = lstCons.SelectedItem.SubItems(1)
    txtModel = lstCons.SelectedItem.SubItems(2)
    txtEngine = lstCons.SelectedItem.SubItems(3)
    txtPlateno = lstCons.SelectedItem.SubItems(4)
    txtColor = lstCons.SelectedItem.SubItems(5)
    txtSerial = lstCons.SelectedItem.SubItems(6)
    txtprdn = lstCons.SelectedItem.SubItems(7)
    txtTIN = lstCons.SelectedItem.SubItems(8)
    txtWCN = lstCons.SelectedItem.SubItems(9)
    txtKMR = lstCons.SelectedItem.SubItems(10)
    txtVIN = lstCons.SelectedItem.SubItems(11)
    txtConduction = lstCons.SelectedItem.SubItems(12)
    txtdateSold = lstCons.SelectedItem.SubItems(13)
    txtDateDel = lstCons.SelectedItem.SubItems(14)
    labID.Caption = lstCons.SelectedItem.SubItems(15)
    labVechicle.Caption = Trim(txtYear) & "  " & Trim(txtMake) & "  " & Trim(txtModel) & "  " & Trim(txtEngine)
    Frame1.Enabled = False
End Sub
