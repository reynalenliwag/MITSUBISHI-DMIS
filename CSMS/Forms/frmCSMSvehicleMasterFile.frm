VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCSMSvehicleMasterFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Master File"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9630
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C0C0FF&
      Height          =   2385
      Left            =   3750
      ScaleHeight     =   2325
      ScaleWidth      =   4965
      TabIndex        =   63
      Top             =   2220
      Width           =   5025
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2265
         Left            =   30
         ScaleHeight     =   2235
         ScaleWidth      =   4875
         TabIndex        =   64
         Top             =   30
         Width           =   4905
      End
   End
   Begin VB.Frame fmeCusInfo 
      Caption         =   "Customer Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   9375
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search Customer"
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
         Left            =   4860
         MouseIcon       =   "frmCSMSvehicleMasterFile.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSvehicleMasterFile.frx":0152
         TabIndex        =   15
         ToolTipText     =   "Cancel"
         Top             =   1350
         Width           =   2115
      End
      Begin VB.Label lblCusInfo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   1740
         TabIndex        =   9
         Top             =   1350
         Width           =   2955
      End
      Begin VB.Label lblCusInfo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   1740
         TabIndex        =   8
         Top             =   990
         Width           =   7395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1560
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   2
         Left            =   735
         TabIndex        =   6
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   390
         Width           =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CustomerName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   345
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblCusInfo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   0
         Left            =   1740
         TabIndex        =   3
         Top             =   270
         Width           =   885
      End
      Begin VB.Label lblCusInfo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   1740
         TabIndex        =   2
         Top             =   630
         Width           =   7395
      End
   End
   Begin VB.Frame fmeVehInfo 
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
      Height          =   6285
      Left            =   120
      TabIndex        =   0
      Top             =   1830
      Width           =   9375
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   1050
         ScaleHeight     =   855
         ScaleWidth      =   6075
         TabIndex        =   57
         Top             =   5280
         Width           =   6075
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
            MouseIcon       =   "frmCSMSvehicleMasterFile.frx":0490
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSvehicleMasterFile.frx":05E2
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Add Vehicle"
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
            MouseIcon       =   "frmCSMSvehicleMasterFile.frx":08F5
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSvehicleMasterFile.frx":0A47
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Edit Selected Vehicle"
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
            MouseIcon       =   "frmCSMSvehicleMasterFile.frx":0DA3
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSvehicleMasterFile.frx":0EF5
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Delete Selected Vehicle"
            Top             =   45
            Width           =   705
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
            Left            =   5355
            MouseIcon       =   "frmCSMSvehicleMasterFile.frx":1220
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSvehicleMasterFile.frx":1372
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Exit Window"
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4995
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   9195
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   2985
            Left            =   5070
            TabIndex        =   34
            Top             =   210
            Width           =   4005
            Begin VB.TextBox txtConduction 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1650
               MaxLength       =   8
               TabIndex        =   42
               Top             =   1830
               Width           =   2265
            End
            Begin VB.TextBox txtDateDel 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1650
               MaxLength       =   10
               TabIndex        =   41
               Top             =   2550
               Width           =   2265
            End
            Begin VB.TextBox txtdateSold 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1650
               MaxLength       =   10
               TabIndex        =   40
               Top             =   2190
               Width           =   2265
            End
            Begin VB.TextBox txtWCN 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1650
               MaxLength       =   15
               TabIndex        =   39
               Top             =   750
               Width           =   2265
            End
            Begin VB.TextBox txtTIN 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1650
               MaxLength       =   15
               TabIndex        =   38
               Top             =   390
               Width           =   2265
            End
            Begin VB.TextBox txtVIN 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1650
               MaxLength       =   20
               TabIndex        =   37
               Top             =   1470
               Width           =   2265
            End
            Begin VB.TextBox txtKMR 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1650
               MaxLength       =   10
               TabIndex        =   36
               Top             =   1110
               Width           =   2265
            End
            Begin VB.TextBox txtprdn 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1650
               MaxLength       =   6
               TabIndex        =   35
               Top             =   30
               Width           =   2265
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   14
               Left            =   300
               TabIndex        =   50
               Top             =   1920
               Width           =   1275
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   12
               Left            =   330
               TabIndex        =   49
               Top             =   2640
               Width           =   1230
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   11
               Left            =   720
               TabIndex        =   48
               Top             =   2310
               Width           =   810
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   10
               Left            =   30
               TabIndex        =   47
               Top             =   840
               Width           =   1560
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   9
               Left            =   960
               TabIndex        =   46
               Top             =   480
               Width           =   585
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   8
               Left            =   960
               TabIndex        =   45
               Top             =   1560
               Width           =   600
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   7
               Left            =   0
               TabIndex        =   44
               Top             =   1200
               Width           =   1560
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   6
               Left            =   330
               TabIndex        =   43
               Top             =   120
               Width           =   1245
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   3375
            Left            =   90
            TabIndex        =   16
            Top             =   210
            Width           =   4725
            Begin VB.TextBox txtColor 
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
               Height          =   330
               Left            =   900
               MaxLength       =   4
               TabIndex        =   62
               Top             =   1950
               Width           =   2475
            End
            Begin VB.CommandButton cmdNewColor 
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
               TabIndex        =   25
               Top             =   1950
               Width           =   345
            End
            Begin VB.TextBox txtSerial 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   900
               MaxLength       =   18
               TabIndex        =   24
               Top             =   2310
               Width           =   2475
            End
            Begin VB.TextBox txtPlateno 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   900
               MaxLength       =   6
               TabIndex        =   23
               Top             =   1560
               Width           =   2475
            End
            Begin VB.TextBox txtEngine 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   900
               TabIndex        =   22
               Top             =   1170
               Width           =   2475
            End
            Begin VB.TextBox txtyear 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   900
               MaxLength       =   4
               TabIndex        =   21
               Top             =   0
               Width           =   975
            End
            Begin VB.ComboBox cboSellingDealer 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   900
               TabIndex        =   20
               Text            =   "cboSellingDealer"
               Top             =   2730
               Width           =   2535
            End
            Begin VB.CommandButton Command1 
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
               Left            =   3450
               TabIndex        =   19
               Top             =   2730
               Width           =   345
            End
            Begin VB.ComboBox cboMake 
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
               Left            =   900
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   390
               Width           =   2505
            End
            Begin VB.ComboBox cboModel 
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
               Left            =   900
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   780
               Width           =   2505
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   13
               Left            =   360
               TabIndex        =   33
               Top             =   1980
               Width           =   450
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Serial / Engine#"
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
               Height          =   465
               Index           =   5
               Left            =   150
               TabIndex        =   32
               Top             =   2250
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   4
               Left            =   60
               TabIndex        =   31
               Top             =   1650
               Width           =   750
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   3
               Left            =   240
               TabIndex        =   30
               Top             =   1260
               Width           =   570
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   2
               Left            =   300
               TabIndex        =   29
               Top             =   900
               Width           =   510
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   1
               Left            =   360
               TabIndex        =   28
               Top             =   480
               Width           =   465
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00800000&
               Height          =   225
               Index           =   0
               Left            =   420
               TabIndex        =   27
               Top             =   90
               Width           =   390
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Selling Dealer"
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
               Height          =   465
               Index           =   16
               Left            =   210
               TabIndex        =   26
               Top             =   2820
               Width           =   900
            End
         End
         Begin MSComctlLib.ListView lsvVEH 
            Height          =   1245
            Left            =   90
            TabIndex        =   14
            Top             =   3600
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   2196
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
            Appearance      =   0
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
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Make"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Engine"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "PlateNo"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Color"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Serial no"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Selling Dealer"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Prod. No."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Tin no."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Warranty Cer."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "KM Rdg"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "VIN No"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "Conductn No."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "Date Sold"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Text            =   "Date Delivered"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   7680
         ScaleHeight     =   885
         ScaleWidth      =   1590
         TabIndex        =   10
         Top             =   5310
         Width           =   1590
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
            MouseIcon       =   "frmCSMSvehicleMasterFile.frx":16D8
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSvehicleMasterFile.frx":182A
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Save New Vehicle"
            Top             =   30
            Width           =   705
         End
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
            MouseIcon       =   "frmCSMSvehicleMasterFile.frx":1B7A
            MousePointer    =   99  'Custom
            Picture         =   "frmCSMSvehicleMasterFile.frx":1CCC
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Cancel"
            Top             =   30
            Width           =   705
         End
      End
   End
   Begin VB.PictureBox picSearchCus 
      BackColor       =   &H00C0C0FF&
      Height          =   3165
      Left            =   2760
      ScaleHeight     =   3105
      ScaleWidth      =   4305
      TabIndex        =   51
      Top             =   1410
      Visible         =   0   'False
      Width           =   4365
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3045
         Left            =   30
         ScaleHeight     =   3015
         ScaleWidth      =   4215
         TabIndex        =   52
         Top             =   30
         Width           =   4245
         Begin VB.CommandButton cmdExitSearch 
            BackColor       =   &H000000FF&
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
            Height          =   255
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   0
            Width           =   255
         End
         Begin VB.TextBox txtSearchCus 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   60
            TabIndex        =   54
            Top             =   390
            Width           =   4035
         End
         Begin MSComctlLib.ListView lsvCus 
            Height          =   2055
            Left            =   90
            TabIndex        =   53
            Top             =   840
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Customer Name"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Code"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search Customer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   55
            Top             =   120
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmCSMSvehicleMasterFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboMake_Change()
    Call FillModel
End Sub

Sub FillModel()
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = gconDMIS.Execute("Select Distinct Model From CSMS_MODELS Where Make = '" & cboMake.Text & "' ORder By MOdel")
    cboModel.Clear
    
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            cboModel.AddItem Null2String(rsTmp!Model)
            
            rsTmp.MoveNext
        Loop
    End If
    
    Set rsTmp = Nothing
End Sub

Private Sub cboMake_Click()
    Call FillModel
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExitSearch_Click()
    Call DisableFrame(True)
End Sub

Private Sub cmdNewColor_Click()
    frmCSMSGetColor.Show 1
End Sub

Private Sub cmdSearch_Click()
    Call DisableFrame(False)

    picSearchCus.Visible = True
    picSearchCus.ZOrder 0
    txtSearchCus.Text = ""
    txtSearchCus.SetFocus
End Sub

Sub DisableFrame(COND As Boolean)
    fmeCusInfo.Enabled = COND
    fmeVehInfo.Enabled = COND
End Sub

Private Sub Command1_Click()
    Dim dsellingdealer As String
    
    dsellingdealer = cboSellingDealer

    frmCSMS_Files_SellingDealer.Show 1
    
    Combo_Loadval cboSellingDealer, gconDMIS.Execute("SELECT SELLINGdealer FROM CSMS_SellingDealer")
    cboSellingDealer = dsellingdealer
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    Call FillMake
    Call FillSellingDealer
End Sub

Sub FillSellingDealer()
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = gconDMIS.Execute("Select Distinct SellingDealer From CSMS_SellingDealer ORder By SellingDealer")
    cboSellingDealer.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            cboSellingDealer.AddItem Null2String(rsTmp!sellingdealer)
                     
            rsTmp.MoveNext
        Loop
    End If
    
    Set rsTmp = Nothing
End Sub

Sub FillMake()
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = gconDMIS.Execute("Select Distinct Make From CSMS_Models ORder By Make")
    cboMake.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            cboMake.AddItem Null2String(rsTmp!Make)
                     
            rsTmp.MoveNext
        Loop
    End If
    
    Set rsTmp = Nothing
End Sub

Private Sub lsvCus_DblClick()
    Dim Index As Double
    
    If Not lsvCus.ListItems.Count = 0 Then
        Index = lsvCus.SelectedItem.Index
        
        With lsvCus
            lblCusInfo(0).Caption = .ListItems(Index).SubItems(3)
            lblCusInfo(1).Caption = .ListItems(Index).Text
            lblCusInfo(2).Caption = .ListItems(Index).SubItems(1)
            lblCusInfo(3).Caption = .ListItems(Index).SubItems(2)
            
            picSearchCus.Visible = False
            Call DisableFrame(True)
            
            Call FillVehicle
        End With
    End If
End Sub

Function GetSellingDealer(xxx)
    Dim temprs                          As ADODB.Recordset
    Set temprs = gconDMIS.Execute("select sellingDealer from CSMS_SellingDealer where sellingCode='" & xxx & "'")
    If Not temprs.BOF Or Not temprs.EOF Then
        GetSellingDealer = Null2String(temprs!sellingdealer)
    End If
    Set temprs = Nothing

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

Sub FillVehicle()
    Dim rsTmp As New ADODB.Recordset
    Dim ITEM As ListItem
    
    Set rsTmp = gconDMIS.Execute("Select * From CSMS_CusVeh Where CusCde = '" & lblCusInfo(0).Caption & "'")
    lsvVEH.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvVEH.ListItems.Add(, , Null2String(rsTmp!yER))
            ITEM.SubItems(1) = Null2String(rsTmp!Make)
            ITEM.SubItems(2) = Null2String(rsTmp!Model)
            ITEM.SubItems(3) = Null2String(rsTmp!Engine)
            ITEM.SubItems(4) = Null2String(rsTmp!Plate_No)
            ITEM.SubItems(5) = GetColor(Null2String(rsTmp!CLRCDE))
            ITEM.SubItems(6) = Null2String(rsTmp!Serial)
            ITEM.SubItems(7) = GetSellingDealer(Null2String(rsTmp!selling_dealer))
            ITEM.SubItems(8) = Null2String(rsTmp!ProdNo)
            ITEM.SubItems(9) = Null2String(rsTmp!tin_number)
            ITEM.SubItems(10) = Null2String(rsTmp!WAR_CERT)
            ITEM.SubItems(11) = Null2String(rsTmp!KMReading)
            ITEM.SubItems(12) = Null2String(rsTmp!VIN)
            ITEM.SubItems(13) = Null2String(rsTmp!vcond_no)
            ITEM.SubItems(14) = Null2String(rsTmp!D_Sold)
            ITEM.SubItems(15) = Null2String(rsTmp!del_date)
            
            rsTmp.MoveNext
        Loop
    End If
    
    Set rsTmp = Nothing
End Sub

Private Sub lsvVEH_Click()
    Dim Index As Double
    
    If Not lsvCus.ListItems.Count = 0 Then
        Index = lsvCus.SelectedItem.Index
        
        With lsvCus
                        
        
        End With
    End If
End Sub

Private Sub txtSearchCus_Change()
    Dim rsTmp As New ADODB.Recordset
    Dim Keyword As String
    Dim ITEM As ListItem
    
    Keyword = txtSearchCus.Text
    
    If Not txtSearchCus.Text = "" Then
        Set rsTmp = gconDMIS.Execute("Select AcctName,Cuscde,CustomerADd,HomePhone,TelephoneNo From ALL_CUSTOMER Where AcctName Like '" & Keyword & "%' ORder By AcctName ASC")
        
        lsvCus.ListItems.Clear
        If Not (rsTmp.BOF And rsTmp.EOF) Then
            Do While Not rsTmp.EOF
                Set ITEM = lsvCus.ListItems.Add(, , Null2String(rsTmp!acctname))
                ITEM.SubItems(1) = Null2String(rsTmp!CustomerAdd)
                ITEM.SubItems(2) = Null2String(rsTmp!HomePhone) & " / " & Null2String(rsTmp!TelephoneNo)
                ITEM.SubItems(3) = Null2String(rsTmp!CUSCDE)
                    
                rsTmp.MoveNext
            Loop
        Else
            lsvCus.ListItems.Clear
        End If
    Else
        lsvCus.ListItems.Clear
    End If
    
    Set rsTmp = Nothing
End Sub
