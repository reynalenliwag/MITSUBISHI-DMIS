VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSYrMkMlEgn 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
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
   Icon            =   "YrMkMlEgn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5220
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   9060
      ScaleHeight     =   3855
      ScaleWidth      =   5475
      TabIndex        =   37
      Top             =   30
      Visible         =   0   'False
      Width           =   5475
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
         Left            =   1470
         MouseIcon       =   "YrMkMlEgn.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "YrMkMlEgn.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Delete Selected Record"
         Top             =   1650
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
         Left            =   750
         MouseIcon       =   "YrMkMlEgn.frx":0D47
         MousePointer    =   99  'Custom
         Picture         =   "YrMkMlEgn.frx":0E99
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Edit Selected Record"
         Top             =   1650
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1230
         TabIndex        =   42
         Top             =   180
         Width           =   2925
      End
      Begin VB.CommandButton Command2 
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
         Height          =   675
         Left            =   3510
         MouseIcon       =   "YrMkMlEgn.frx":11F5
         MousePointer    =   99  'Custom
         Picture         =   "YrMkMlEgn.frx":1347
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1590
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2820
         MouseIcon       =   "YrMkMlEgn.frx":1685
         MousePointer    =   99  'Custom
         Picture         =   "YrMkMlEgn.frx":17D7
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1590
         Width           =   645
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1230
         TabIndex        =   39
         Top             =   600
         Width           =   2925
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1260
         TabIndex        =   38
         Top             =   1050
         Width           =   2925
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Model Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   45
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   660
         Width           =   885
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&XXX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   43
         Top             =   1080
         Width           =   885
      End
   End
   Begin VB.PictureBox picEngine 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   9060
      ScaleHeight     =   3855
      ScaleWidth      =   5475
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   5475
      Begin VB.CommandButton cmdSaveEngine 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2280
         MouseIcon       =   "YrMkMlEgn.frx":1B27
         MousePointer    =   99  'Custom
         Picture         =   "YrMkMlEgn.frx":1C79
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2910
         Width           =   645
      End
      Begin VB.CommandButton cmdCancelEngine 
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
         Height          =   675
         Left            =   2970
         MouseIcon       =   "YrMkMlEgn.frx":1FC9
         MousePointer    =   99  'Custom
         Picture         =   "YrMkMlEgn.frx":211B
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2910
         Width           =   645
      End
      Begin VB.TextBox txtEngineVIN 
         Height          =   330
         Left            =   2280
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2490
         Width           =   1245
      End
      Begin VB.ComboBox cboEngineAspiration 
         Height          =   345
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2105
         Width           =   1335
      End
      Begin VB.TextBox txtEngineFuelType 
         Height          =   330
         Left            =   2280
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1735
         Width           =   1245
      End
      Begin VB.TextBox txtEngineDisplacement 
         Height          =   330
         Left            =   2280
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1365
         Width           =   1245
      End
      Begin VB.TextBox txtEngineCubic 
         Height          =   330
         Left            =   2280
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   995
         Width           =   1245
      End
      Begin VB.TextBox txtEngineLiters 
         Height          =   330
         Left            =   2280
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   625
         Width           =   1245
      End
      Begin VB.TextBox txtEnginetype 
         Height          =   330
         Left            =   2280
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   270
         Width           =   2325
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine VIN"
         Height          =   225
         Left            =   1245
         TabIndex        =   17
         Top             =   2520
         Width           =   915
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Aspiration"
         Height          =   225
         Left            =   1335
         TabIndex        =   16
         Top             =   2145
         Width           =   825
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Type"
         Height          =   225
         Left            =   1365
         TabIndex        =   15
         Top             =   1785
         Width           =   795
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Cubic Inch Displacement"
         Height          =   225
         Left            =   90
         TabIndex        =   14
         Top             =   1410
         Width           =   2070
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Cubic Centimeters"
         Height          =   225
         Left            =   615
         TabIndex        =   13
         Top             =   1035
         Width           =   1545
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Liters"
         Height          =   225
         Left            =   1695
         TabIndex        =   12
         Top             =   675
         Width           =   465
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine Type"
         Height          =   225
         Left            =   1140
         TabIndex        =   11
         Top             =   300
         Width           =   1020
      End
   End
   Begin VB.PictureBox picMake 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   9060
      ScaleHeight     =   3855
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   5475
      Begin VB.TextBox txtMakeFlatRate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1260
         TabIndex        =   36
         Top             =   1050
         Width           =   2925
      End
      Begin VB.TextBox txtMake 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1230
         TabIndex        =   22
         Top             =   600
         Width           =   2925
      End
      Begin VB.CommandButton cmdMakeCancel 
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
         Height          =   675
         Left            =   3510
         MouseIcon       =   "YrMkMlEgn.frx":2459
         MousePointer    =   99  'Custom
         Picture         =   "YrMkMlEgn.frx":25AB
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1590
         Width           =   645
      End
      Begin VB.TextBox txtMakeCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1230
         TabIndex        =   2
         Top             =   180
         Width           =   2925
      End
      Begin VB.CommandButton cmdMakeSave 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2880
         MouseIcon       =   "YrMkMlEgn.frx":28E9
         MousePointer    =   99  'Custom
         Picture         =   "YrMkMlEgn.frx":2A3B
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1590
         Width           =   645
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Flat Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   35
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Make"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   660
         Width           =   885
      End
      Begin VB.Label labMakeCode 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   0
      ScaleHeight     =   3765
      ScaleWidth      =   5175
      TabIndex        =   27
      Top             =   0
      Width           =   5205
      Begin VB.CommandButton Command4 
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
         Left            =   4500
         MouseIcon       =   "YrMkMlEgn.frx":2D8B
         MousePointer    =   99  'Custom
         Picture         =   "YrMkMlEgn.frx":2EDD
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Add Record"
         Top             =   330
         Width           =   435
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   2295
         Width           =   3345
      End
      Begin VB.CommandButton Command3 
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
         Left            =   4500
         MouseIcon       =   "YrMkMlEgn.frx":3047
         MousePointer    =   99  'Custom
         Picture         =   "YrMkMlEgn.frx":3199
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Add Record"
         Top             =   1680
         Width           =   435
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   945
         Left            =   1440
         ScaleHeight     =   945
         ScaleWidth      =   4485
         TabIndex        =   33
         Top             =   2820
         Width           =   4485
         Begin VB.CommandButton cmdCancel 
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
            Left            =   2760
            MouseIcon       =   "YrMkMlEgn.frx":3303
            MousePointer    =   99  'Custom
            Picture         =   "YrMkMlEgn.frx":3455
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Exit Window"
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "&Select"
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
            Left            =   2040
            MouseIcon       =   "YrMkMlEgn.frx":37BB
            MousePointer    =   99  'Custom
            Picture         =   "YrMkMlEgn.frx":390D
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Select"
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.ComboBox cboEngine 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   4170
         Width           =   2625
      End
      Begin VB.ComboBox cboModel 
         Height          =   345
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1695
         Width           =   2715
      End
      Begin VB.ComboBox cboMake 
         Height          =   345
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1050
         Width           =   2625
      End
      Begin VB.ComboBox cboYear 
         Height          =   345
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   345
         Width           =   2715
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   4965
         Left            =   0
         ScaleHeight     =   4965
         ScaleWidth      =   1725
         TabIndex        =   28
         Top             =   -30
         Width           =   1725
         Begin VB.Label labformname 
            Height          =   255
            Left            =   390
            TabIndex        =   72
            Top             =   2280
            Width           =   915
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "MAKE"
            Height          =   225
            Left            =   180
            MouseIcon       =   "YrMkMlEgn.frx":3C49
            MousePointer    =   99  'Custom
            TabIndex        =   32
            Top             =   870
            Width           =   825
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   345
            Left            =   90
            Top             =   810
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            Height          =   225
            Left            =   150
            MouseIcon       =   "YrMkMlEgn.frx":3F53
            MousePointer    =   99  'Custom
            TabIndex        =   31
            Top             =   1710
            Width           =   1305
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            Height          =   225
            Left            =   180
            MouseIcon       =   "YrMkMlEgn.frx":425D
            MousePointer    =   99  'Custom
            TabIndex        =   30
            Top             =   1290
            Width           =   795
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "YEAR"
            Height          =   195
            Left            =   180
            MouseIcon       =   "YrMkMlEgn.frx":4567
            MousePointer    =   99  'Custom
            TabIndex        =   29
            Top             =   420
            Width           =   675
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   345
            Left            =   90
            Top             =   360
            Width           =   1515
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   345
            Left            =   90
            Top             =   1230
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   345
            Left            =   90
            Top             =   1650
            Visible         =   0   'False
            Width           =   1515
         End
      End
      Begin VB.Label Label22 
         Caption         =   "MODEL DESCRIPTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1770
         TabIndex        =   57
         Top             =   2070
         Width           =   2625
      End
      Begin VB.Label Label5 
         Caption         =   "YEAR :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1770
         TabIndex        =   53
         Top             =   90
         Width           =   2625
      End
      Begin VB.Label Label6 
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
         Height          =   315
         Left            =   1770
         TabIndex        =   52
         Top             =   795
         Width           =   2625
      End
      Begin VB.Label Label7 
         Caption         =   "ENGINE"
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
         TabIndex        =   51
         Top             =   3870
         Width           =   2625
      End
      Begin VB.Label Label8 
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
         Height          =   285
         Left            =   1770
         TabIndex        =   50
         Top             =   1470
         Width           =   2625
      End
   End
   Begin VB.PictureBox picYear 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   5205
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   5205
      Begin MSComctlLib.ListView ListView1 
         Height          =   2325
         Left            =   60
         TabIndex        =   61
         Top             =   540
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4101
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtYear 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   840
         TabIndex        =   25
         Top             =   120
         Width           =   2925
      End
      Begin VB.PictureBox picSaveYear 
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   2820
         ScaleHeight     =   825
         ScaleWidth      =   2445
         TabIndex        =   67
         Top             =   2910
         Visible         =   0   'False
         Width           =   2445
         Begin VB.CommandButton cmdYearCancel 
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
            Left            =   1530
            MouseIcon       =   "YrMkMlEgn.frx":4871
            MousePointer    =   99  'Custom
            Picture         =   "YrMkMlEgn.frx":49C3
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   0
            Width           =   765
         End
         Begin VB.CommandButton cmdYearSave 
            Caption         =   "&OK"
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
            Left            =   780
            MouseIcon       =   "YrMkMlEgn.frx":4D01
            MousePointer    =   99  'Custom
            Picture         =   "YrMkMlEgn.frx":4E53
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.PictureBox picAddYear 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   300
         ScaleHeight     =   885
         ScaleWidth      =   6195
         TabIndex        =   62
         Top             =   2880
         Width           =   6195
         Begin VB.CommandButton Command7 
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
            Left            =   4020
            MouseIcon       =   "YrMkMlEgn.frx":51A3
            MousePointer    =   99  'Custom
            Picture         =   "YrMkMlEgn.frx":52F5
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   30
            Width           =   765
         End
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
            Left            =   5415
            MouseIcon       =   "YrMkMlEgn.frx":5633
            MousePointer    =   99  'Custom
            Picture         =   "YrMkMlEgn.frx":5785
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Exit Window"
            Top             =   45
            Width           =   765
         End
         Begin VB.CommandButton cmdDeleteYear 
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
            Left            =   3270
            MouseIcon       =   "YrMkMlEgn.frx":5AEB
            MousePointer    =   99  'Custom
            Picture         =   "YrMkMlEgn.frx":5C3D
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Delete Selected Record"
            Top             =   30
            Width           =   765
         End
         Begin VB.CommandButton cmdEditYear 
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
            Left            =   2520
            MouseIcon       =   "YrMkMlEgn.frx":5F68
            MousePointer    =   99  'Custom
            Picture         =   "YrMkMlEgn.frx":60BA
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Edit Selected Record"
            Top             =   30
            Width           =   765
         End
         Begin VB.CommandButton cmdAddYear 
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
            Left            =   1770
            MouseIcon       =   "YrMkMlEgn.frx":6416
            MousePointer    =   99  'Custom
            Picture         =   "YrMkMlEgn.frx":6568
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Add Record"
            Top             =   30
            Width           =   765
         End
      End
      Begin VB.Label labyearID 
         Caption         =   "Label23"
         Height          =   375
         Left            =   3810
         TabIndex        =   70
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -240
         TabIndex        =   26
         Top             =   210
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmCSMSYrMkMlEgn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLoad                              As ADODB.Recordset
Dim xENGINE                             As String
Dim xLiters                             As String
Dim xCubic                              As String
Dim xDisplacement                       As String
Dim xFuelType                           As String
Dim xAspiration                         As String
Dim xEngineVIN                          As String
Dim AddorEdit                           As String
Dim rsYear                              As ADODB.Recordset
Public DefaultYear                      As String
Public DefaultMake                      As String
Public DefaultModel                     As String
Public DefaultEngine                    As String
Event frmSelectMakeMode()

Private Sub cboEngine_GotFocus()
    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
    Shape4.Visible = True
End Sub

Private Sub cboMake_Click()
    LoadModel
End Sub

Private Sub cboMake_GotFocus()
    Shape1.Visible = False
    Shape2.Visible = True
    Shape3.Visible = False
    Shape4.Visible = False
End Sub

Private Sub cboModel_Click()
    LoadDescription
End Sub

Private Sub cboModel_GotFocus()
    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = True
    Shape4.Visible = False
End Sub



Private Sub cboYear_GotFocus()
    Shape1.Visible = True
    Shape2.Visible = False
    Shape3.Visible = False
    Shape4.Visible = False
End Sub


Private Sub cmdAddYear_Click()
    txtyear = "": AddorEdit = "ADD": ListView1.Enabled = False: txtyear.Enabled = True: picSaveYear.Visible = True: picAddYear.Visible = False
    On Error Resume Next
    txtyear.SetFocus

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelEngine_Click()

    On Error GoTo ERRORCODE:

    picEngine.Visible = False
    picEngine.ZOrder 1

    Exit Sub
ERRORCODE:
    ShowVBError

End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ERRORCODE:





    Exit Sub
ERRORCODE:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()

    On Error GoTo ERRORCODE:

    AddorEdit = "EDIT"
    If Shape1.Visible = True Then
        picYear.Visible = True
        picYear.ZOrder 0
        txtyear = cboYear.Text
        On Error Resume Next
        txtyear.SetFocus
    ElseIf Shape2.Visible = True Then
        picMake.Visible = True
        picMake.ZOrder 0
        FillMake
        On Error Resume Next
        txtMakeCode.SetFocus
        picEngine.Visible = False
    ElseIf Shape3.Visible = True Then

    ElseIf Shape4.Visible = True Then
        picEngine.Visible = True
        picEngine.ZOrder 0
    End If





    Exit Sub
ERRORCODE:
    ShowVBError
End Sub

Private Sub cmdMakeCancel_Click()
    picMake.Visible = False
End Sub

Private Sub cmdMakeSave_Click()

    On Error GoTo ERRORCODE

    If LTrim(RTrim(txtMakeCode)) = "" Then
        ShowIsRequiredMsg " MAKE CODE"
        On Error Resume Next
        txtMakeCode.SetFocus
        Exit Sub
    End If

    If LTrim(RTrim(txtMake)) = "" Then
        ShowIsRequiredMsg " MAKE "
        On Error Resume Next
        txtMake.SetFocus
        Exit Sub
    End If



    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into ALL_MAKE " & _
                       " (CODE,MAKE, FLATRATE)" & _
                       " values(" & N2Str2Null(txtMakeCode) & ", " & N2Str2Null(txtMake) & "," & NumericVal(txtMakeFlatRate) & ")"
    Else
        If cboMake.ListIndex = -1 Then
            Exit Sub
        End If

        gconDMIS.Execute "UPDATE ALL_MAKE SET " & _
                       "  CODE=" & N2Str2Null(txtMakeCode) & "," & _
                       "  MAKE=" & N2Str2Null(txtMake) & "," & _
                       "  FLATRATE=" & NumericVal(txtMakeFlatRate) & _
                       " Where ID=" & cboMake.ItemData(cboMake.ListIndex)



    End If

    LoadMake
    If cboMake.ListCount > 0 Then
        cboMake.ListIndex = 0
    End If
    cmdMakeCancel.Value = True

    Exit Sub

ERRORCODE:

    ShowVBError
    Exit Sub
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo ERRORCODE:
    Select Case labformname
        Case "frmCSMSAddVehicle"
            With frmCSMSAddVehicle
                .txtyear = cboYear
                .txtMake = cboMake
                .txtModel = cboModel
                .txtEngine = cboEngine
                .txtDescription = cboYear & " " & Combo1
                Unload Me
            End With
            Case "frmCSMSEditAppVehicle"
            With frmCSMSEditAppVehicle
                .txtyear = cboYear
                .txtMake = cboMake
                .txtModel = cboModel
                .txtEngine = cboEngine
                .txtDescription = cboYear & " " & Combo1
            End With
            Unload Me
        End Select
        Exit Sub
ERRORCODE:
        ShowVBError
    End Sub

Private Sub cmdYearCancel_Click()
    ListView1.Enabled = True
    txtyear.Enabled = False
    picSaveYear.Visible = False
    picAddYear.Visible = True
    StoreMemVarsYear
End Sub

Private Sub cmdYearSave_Click()
    On Error GoTo ERRORCODE:
    Dim checkRs                         As ADODB.Recordset
    If LTrim(RTrim(txtyear)) = "" Then
        ShowIsRequiredMsg " YEAR"
        On Error Resume Next
        txtyear.SetFocus
        Exit Sub
    End If
    Set checkRs = gconDMIS.Execute("Select count(*) from all_year where yeer='" & LTrim(RTrim(txtyear)) & "'")
    If AddorEdit = "ADD" Then
        If checkRs(0).Value > 0 Then
            MsgBox "Such Year Already Exists ", vbInformation
            Exit Sub
        End If
    Else
        If Null2String(rsYear!YEER) <> LTrim(RTrim(txtyear)) Then
            If checkRs(0).Value > 0 Then
                MsgBox "Such Year Already Exists", vbInformation
                Exit Sub
            End If
        End If
    End If

    If AddorEdit = "ADD" Then
        MsgBox "New Year Has Been Added", vbInformation

        gconDMIS.Execute "Insert into All_Year " & _
                       " (yeer)" & _
                       " values(" & txtyear & ")"
    Else
        If cboYear.ItemData(cboYear.ListIndex) = -1 Then
            Exit Sub
        End If
        gconDMIS.Execute "UPDATE All_Year SET " & _
                       "  yeer=" & Null2String(txtyear) & " Where ID=" & cboYear.ItemData(cboYear.ListIndex)
        MsgBox "Year Has Been Updated", vbInformation

    End If
    rsRefreshYear
    cmdYearCancel.Value = True


    Exit Sub
ERRORCODE:
    ShowVBError

End Sub

Public Sub Combo_LoadList(WeirdoCombo As ComboBox, RecSet As ADODB.Recordset, ItemDataRow As Integer)
    WeirdoCombo.Clear
    If Not (RecSet.BOF And RecSet.EOF) Then
        While Not RecSet.EOF
            WeirdoCombo.AddItem Null2String(RecSet(0))
            If ItemDataRow <> -1 Then
                WeirdoCombo.ItemData(WeirdoCombo.NewIndex) = Null2String(RecSet(ItemDataRow))
            End If
            RecSet.MoveNext
        Wend
    End If
    Set RecSet = Nothing
End Sub



Private Sub cmdSaveEngine_Click()

    On Error GoTo ERRORCODE:

    If LTrim(RTrim(txtEnginetype)) = "" Then
        ShowIsRequiredMsg " Engine Type"
        On Error Resume Next
        txtEnginetype.SetFocus
        Exit Sub
    End If
    If LTrim(RTrim(txtEngineLiters)) = "" Then
        ShowIsRequiredMsg " Engine Liters"
        On Error Resume Next
        txtEngineLiters.SetFocus
        Exit Sub
    End If

    If LTrim(RTrim(txtEngineVIN)) = "" Then
        ShowIsRequiredMsg " Engine VIN"
        On Error Resume Next
        txtEngineVIN.SetFocus
        Exit Sub
    End If

    If LTrim(RTrim(txtEngineFuelType)) = "" Then
        ShowIsRequiredMsg "Fuel Type"
        On Error Resume Next
        txtEngineFuelType.SetFocus
        Exit Sub
    End If

    Dim SQL                             As String






    If txtEnginetype.Text <> "" Then
        xENGINE = N2Str2Null(txtEnginetype)
        xLiters = N2Str2Null(txtEngineLiters)
        xCubic = N2Str2Null(txtEngineCubic)
        xDisplacement = N2Str2Null(txtEngineDisplacement)
        xFuelType = N2Str2Null(txtEngineFuelType)
        xAspiration = N2Str2Null(cboEngineAspiration.Text)
        xEngineVIN = N2Str2Null(txtEngineVIN)

        gconDMIS.Execute "Insert into All_Engine " & _
                       " (ENGINE,Liters,Cubic,Displacement,FuelType,Aspiration,EngineVIN)" & _
                       " values(" & xENGINE & "," & xLiters & "," & xCubic & "," & xDisplacement & "," & xFuelType & "," & xAspiration & "," & xEngineVIN & ")"
        cmdCancelEngine.Value = True
    End If

    Exit Sub
ERRORCODE:
    ShowVBError

End Sub


Sub FillMake()
    If cboMake.ListIndex = -1 Then Exit Sub
    Dim temprs                          As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT * FROM ALL_MAKE WHERE ID=" & cboMake.ItemData(cboMake.ListIndex) & " Order by ID ASC")
    If Not (temprs.EOF Or temprs.BOF) Then
        txtMakeCode = Null2String(temprs!code)
        txtMake = Null2String(temprs!Make)
        txtMakeFlatRate = NumericVal(temprs!FLATRATE)
    End If

End Sub

Private Sub Combo1_GotFocus()
    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
    Shape4.Visible = True

End Sub

Private Sub Command3_Click()
    frmCSMSModel.Show 1
    LoadModel
End Sub

Private Sub Command4_Click()
    rsRefreshYear
    StoreMemVarsYear
    picYear.Visible = True
    picYear.ZOrder 0


End Sub

Private Sub cmdDeleteYear_Click()
    If MsgBox("Are You Sure You Want to Delete this Record", vbInformation + vbYesNo) = vbYes Then
        gconDMIS.Execute ("Delete from all_year where id=" & labyearID)
        rsRefreshYear
        StoreMemVarsYear
    End If

End Sub
Sub StoreMemVarsYear()
    If Not rsYear.EOF Or Not rsYear.BOF Then
        labyearID = rsYear!ID
        txtyear = Null2String(rsYear!YEER)
    Else
        cmdAddYear.Value = True
    End If
End Sub

Sub rsRefreshYear()
    Set rsYear = New ADODB.Recordset
    Call rsYear.Open("select * from all_year order by yeer", gconDMIS, adOpenDynamic, adLockReadOnly)
    Dim I
    ListView1.ListItems.Clear
    While Not rsYear.EOF
        I = I + 1
        ListView1.ListItems.Add , , Null2String(rsYear!YEER)

        ListView1.ListItems.ITEM(I).ListSubItems.Add , , rsYear!ID
        rsYear.MoveNext
    Wend
    rsYear.MoveFirst
End Sub

Private Sub cmdEditYear_Click()
    AddorEdit = "EDIT": ListView1.Enabled = False: txtyear.Enabled = True: picSaveYear.Visible = True: picAddYear.Visible = False
End Sub

Private Sub Command7_Click()
    LoadYear
    picMain.Visible = True: picMain.ZOrder 0: ListView1.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    InitField
    InitEngine
    LoadYear
    cboYear.ListIndex = GetComboList(cboYear, DefaultYear)
    LoadMake
    cboMake.ListIndex = GetComboList(cboMake, DefaultMake)
    LoadModel
    cboModel.ListIndex = GetComboList(cboModel, DefaultModel)
    LoadEngine
    cboEngine.ListIndex = GetComboList(cboEngine, DefaultEngine)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DefaultYear = ""
    DefaultMake = ""
    DefaultModel = ""
    DefaultEngine = ""
End Sub

Sub InitEngine()
    txtEnginetype = "": txtEngineLiters = "": txtEngineCubic = "": txtEngineDisplacement = ""
    txtEngineFuelType = "": cboEngineAspiration.ListIndex = -1: txtEngineVIN = ""
End Sub

Sub InitField()
    cboYear.Clear
    cboMake.Clear
    cboModel.Clear
    cboEngine.Clear
End Sub

Sub LoadEngine()

    Set rsLoad = New ADODB.Recordset
    Set rsLoad = gconDMIS.Execute("SELECT ENGINE, ID  FROM ALL_Engine order by engine asc")
    cboEngine.Clear
    If Not rsLoad.EOF And Not rsLoad.BOF Then
        Call Combo_Loadval(cboEngine, rsLoad)
    End If


End Sub

Sub LoadMake()

    Set rsLoad = New ADODB.Recordset
    cboMake.Clear
    Set rsLoad = gconDMIS.Execute("Select Make, ID  from All_Make order by ID asc")
    If Not rsLoad.EOF And Not rsLoad.BOF Then
        Combo_LoadList cboMake, rsLoad, 1

    End If

End Sub

Sub LoadModel()
    Set rsLoad = New ADODB.Recordset
    cboModel.Clear
    Set rsLoad = gconDMIS.Execute("Select distinct upper(Model) model  from CSMIOS_S_MODEL WHERE MAKE=" & N2Str2Null(cboMake.Text) & " and isnull(model,'')<>'' order by model asc")
    If Not rsLoad.EOF And Not rsLoad.BOF Then
        Combo_LoadList cboModel, rsLoad, -1
    End If
End Sub

Sub LoadDescription()
    Set rsLoad = New ADODB.Recordset
    Combo1.Clear
    Set rsLoad = gconDMIS.Execute("Select distinct DESCRIPT  from CSMIOS_S_MODEL WHERE model=" & N2Str2Null(cboModel.Text) & " order by 1 asc")
    If Not rsLoad.EOF And Not rsLoad.BOF Then
        Combo_LoadList Combo1, rsLoad, -1
    End If
End Sub

Sub LoadYear()
    Set rsLoad = New ADODB.Recordset
    cboYear.Clear
    Set rsLoad = gconDMIS.Execute("Select yeer,ID from ALL_year order by yeer desc")
    If Not rsLoad.EOF And Not rsLoad.BOF Then
        Combo_LoadList cboYear, rsLoad, 1
    End If
End Sub

Function GetComboList(C As ComboBox, STR As String, Optional ByVal ByItemData As Boolean = False) As Integer
    If C.ListCount = 0 Then: GetComboList = -1: Exit Function
    Dim I                               As Long
    Dim ItemDataX                       As Long
    If ByItemData = False Then
        For I = 0 To C.ListCount - 1
            If UCase(C.List(I)) = UCase(Trim(STR)) Then
                GetComboList = I
                Exit Function
            End If
        Next
    Else
        If STR = vbNullString Then
            GetComboList = -1
            Exit Function
        End If

        ItemDataX = CLng(STR)

        For I = 0 To C.ListCount - 1
            If C.ItemData(I) = STR Then
                GetComboList = I
                Exit Function
            End If
        Next
    End If
    GetComboList = -1
End Function

Private Sub ListView1_ItemClick(ByVal ITEM As MSComctlLib.ListItem)

    rsYear.MoveFirst
    rsYear.Find ("ID=" & ITEM.ListSubItems(1).Text)
    StoreMemVarsYear
End Sub
