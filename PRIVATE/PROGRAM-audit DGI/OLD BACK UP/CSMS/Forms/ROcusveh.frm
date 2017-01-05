VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSROCusveh 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Vehicle Information"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "ROcusveh.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9315
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      Left            =   60
      TabIndex        =   13
      Top             =   -30
      Width           =   9225
      Begin VB.TextBox txtPlateno 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   48
         Top             =   1740
         Width           =   1845
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
         Left            =   5625
         MaxLength       =   100
         TabIndex        =   39
         Top             =   2910
         Width           =   3495
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
         Left            =   5625
         MaxLength       =   50
         TabIndex        =   38
         Top             =   2520
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
         Left            =   5625
         MaxLength       =   100
         TabIndex        =   37
         Top             =   1740
         Width           =   3495
      End
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
         Left            =   5625
         MaxLength       =   50
         TabIndex        =   36
         Top             =   1350
         Width           =   1875
      End
      Begin VB.ComboBox cboEndUser 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1770
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   4500
         Width           =   4095
      End
      Begin VB.ComboBox cboSelling_Dealer 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1770
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   4890
         Width           =   4095
      End
      Begin VB.TextBox txtVCond_No 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   32
         Top             =   2130
         Width           =   1845
      End
      Begin VB.TextBox txtModel 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   960
         Width           =   1845
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   90
         TabIndex        =   12
         Top             =   5640
         Width           =   9015
      End
      Begin VB.TextBox txtMake 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   570
         Width           =   1845
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
         Left            =   2790
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Select Year, Make and Model"
         Top             =   180
         Width           =   345
      End
      Begin VB.TextBox txtyear 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txtWar_Cert 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1770
         MaxLength       =   15
         TabIndex        =   6
         Top             =   3720
         Width           =   1845
      End
      Begin VB.TextBox txtTin_Number 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1770
         MaxLength       =   15
         TabIndex        =   5
         Top             =   3330
         Width           =   1845
      End
      Begin VB.TextBox txtDel_Date 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5640
         MaxLength       =   18
         TabIndex        =   9
         Top             =   960
         Width           =   1845
      End
      Begin VB.TextBox txtD_Sold 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5640
         MaxLength       =   18
         TabIndex        =   8
         Top             =   540
         Width           =   1845
      End
      Begin VB.ComboBox cboColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1770
         TabIndex        =   2
         Top             =   1350
         Width           =   1845
      End
      Begin VB.TextBox txtSerial 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1770
         MaxLength       =   18
         TabIndex        =   7
         Top             =   4110
         Width           =   4065
      End
      Begin VB.TextBox txtEngine 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1770
         MaxLength       =   18
         TabIndex        =   3
         Top             =   2550
         Width           =   1845
      End
      Begin VB.TextBox txtProdNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1770
         MaxLength       =   6
         TabIndex        =   4
         Top             =   2940
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker dtpINS 
         Height          =   345
         Left            =   5625
         TabIndex        =   40
         Top             =   2130
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
         CheckBox        =   -1  'True
         Format          =   16121857
         CurrentDate     =   39647
      End
      Begin MSComCtl2.DTPicker dtpFIN 
         Height          =   345
         Left            =   5625
         TabIndex        =   41
         Top             =   3300
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
         CheckBox        =   -1  'True
         Format          =   16121857
         CurrentDate     =   39647
      End
      Begin VB.Label labid 
         BackColor       =   &H000000FF&
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8520
         TabIndex        =   50
         Top             =   210
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Plate no"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   1005
         TabIndex        =   49
         Top             =   1830
         Width           =   660
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
         Left            =   4260
         TabIndex        =   47
         Top             =   3390
         Width           =   1290
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
         Left            =   4050
         TabIndex        =   46
         Top             =   3000
         Width           =   1500
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
         Left            =   4440
         TabIndex        =   45
         Top             =   2640
         Width           =   1110
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
         Left            =   4260
         TabIndex        =   44
         Top             =   2250
         Width           =   1290
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
         Left            =   3750
         TabIndex        =   43
         Top             =   1830
         Width           =   1800
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
         Left            =   4245
         TabIndex        =   42
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   945
         TabIndex        =   35
         Top             =   4620
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   1080
         TabIndex        =   34
         Top             =   4980
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Conduction Sticker"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   2220
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   1185
         TabIndex        =   31
         Top             =   1050
         Width           =   510
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
         Left            =   90
         TabIndex        =   29
         Top             =   5310
         Width           =   2805
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   1230
         TabIndex        =   28
         Top             =   660
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   1305
         TabIndex        =   27
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Warranty Certificate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   75
         TabIndex        =   20
         Top             =   3810
         Width           =   1635
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TIN Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   750
         TabIndex        =   21
         Top             =   3420
         Width           =   960
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Delivered"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4365
         TabIndex        =   22
         Top             =   1050
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Purchased"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4260
         TabIndex        =   19
         Top             =   660
         Width           =   1290
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "VIN No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1185
         TabIndex        =   18
         Top             =   4200
         Width           =   525
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Engine Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   435
         TabIndex        =   17
         Top             =   2640
         Width           =   1260
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Color Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   765
         TabIndex        =   16
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   14
         Top             =   3030
         Width           =   1350
      End
   End
   Begin VB.PictureBox Picture2 
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
      Height          =   885
      Left            =   7770
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   23
      Top             =   6270
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
         Left            =   720
         MouseIcon       =   "ROcusveh.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "ROcusveh.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Close Window"
         Top             =   60
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
         Left            =   0
         MouseIcon       =   "ROcusveh.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "ROcusveh.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Save Changes"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   630
      TabIndex        =   15
      Top             =   270
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCSMSROCusveh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer                                          As ADODB.Recordset
Dim rsCusVeh                                            As ADODB.Recordset
Dim rsS_Model                                           As ADODB.Recordset
Dim rsColor                                             As ADODB.Recordset
Dim xPLATE_ID                                           As Long
Dim xOLD_PLATENO                                        As String
Dim AddorEdit                                           As String
Dim xFromForm                                           As String
Dim xLOCAL_CUSCDE                                       As String
Dim xLOCAL_ACCTNAME                                     As String
Dim WithEvents frmSelectMakeMode                        As frmCSMSYrMkMlEgn
Attribute frmSelectMakeMode.VB_VarHelpID = -1
Public Event SaveChanges(xPLATE_NO As String, xWARR_CER As String, xMake As String, xMODEL As String, xSERIAL As String, xDESCRIPTION, FromFrom As String)
Public Event SelectionMade(ByVal Code As String, FromForm As String)

Public Sub SelectSQl(XXX As String, FromForm As String, xID As Long, xCuscde As String, xACCTNAME As String, XPLATENO As String)
    Set rsCusVeh = New ADODB.Recordset
    rsCusVeh.Open XXX, gconDMIS, adOpenForwardOnly, adLockReadOnly
    xFromForm = FromForm
    xPLATE_ID = xID
    labID.Caption = xID
    xLOCAL_ACCTNAME = xACCTNAME
    xLOCAL_CUSCDE = xCuscde
    xOLD_PLATENO = XPLATENO
End Sub

Function SetColor(CCC As String)
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select COLOR_CODE,COLOR_DESC from ALL_Color where COLOR_DESC = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        SetColor = Null2String(rsColor!Color_code)
    Else
        SetColor = ""
    End If
End Function

Function SetColorDesc(CCC As String)
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select * from ALL_Color where COLOR_CODE = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        SetColorDesc = Null2String(rsColor!color_desc)
    Else
        SetColorDesc = ""
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
            SetEndUser = Null2String(rsEndUser!ACCTNAME)
        Else
            SetEndUser = Null2String(rsEndUser!CUSCDE)
        End If
    End If
End Function

Sub initMemvars()
    txtYear.Text = ""
    txtMake.Text = ""
    txtModel.Text = ""
    txtPLATENO.Text = ""
    txtDescription.Text = ""
    txtProdNo.Text = ""
    txtModel.Text = ""
    cboCOLOR.Text = ""
    txtENGINE.Text = ""
    txtSerial.Text = ""
    txtD_Sold.Text = ""
    txtWar_Cert.Text = ""
    txtTin_Number.Text = ""
    txtDel_Date.Text = ""
    cboEndUser.Text = ""
    cboSelling_Dealer.Text = ""
    
    Call FillCbo
End Sub

Sub StoreMemVars()
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        labID.Caption = rsCusVeh!ID
        txtYear.Text = Null2String(rsCusVeh!YER)
        txtMake.Text = Null2String(rsCusVeh!Make)
        txtModel.Text = Null2String(rsCusVeh!MODEL)
        txtPLATENO.Text = Null2String(rsCusVeh!PLATE_NO)
        txtDescription.Text = Null2String(rsCusVeh!Description)
        txtProdNo.Text = Null2String(rsCusVeh!ProdNo)
        txtVCond_No.Text = Null2String(rsCusVeh!VCOND_NO)
        txtModel.Text = Null2String(rsCusVeh!MODEL)
        cboCOLOR.Text = SetColorDesc(Null2String(rsCusVeh!ClrCde))
        txtENGINE.Text = Null2String(rsCusVeh!Engine)
        txtSerial.Text = Null2String(rsCusVeh!VIN)
        txtD_Sold.Text = Null2String(rsCusVeh!D_SOLD)
        txtWar_Cert.Text = Null2String(rsCusVeh!War_Cert)
        txtTin_Number.Text = Null2String(rsCusVeh!TIN_Number)
        txtDel_Date.Text = Null2String(rsCusVeh!DEL_DATE)
        cboEndUser.Text = SetEndUser(Null2String(rsCusVeh!EndUser), 1)
        cboSelling_Dealer.Text = SetSellingDealer(Null2String(rsCusVeh!Selling_Dealer), 1)
        AddorEdit = "EDIT"

        'UPDATE BY: MJP 07182008 6:00 PM
            txtINSTYPE.Text = Null2String(rsCusVeh!INS_TYPE)
            txtINSCOMP.Text = Null2String(rsCusVeh!INS_COMP)
            dtpINS.Value = Null2Date(rsCusVeh!INS_EXP_DATE)
            txtFINTYPE.Text = Null2String(rsCusVeh!FIN_TYPE)
            txtFINCOMP.Text = Null2String(rsCusVeh!FIN_COMP)
            dtpFIN.Value = Null2Date(rsCusVeh!FIN_EXP_DATE)
        'UPDATE BY: MJP 07182008 6:00 PM
    Else
        initMemvars
        AddorEdit = "ADD"
    End If
End Sub

Sub FillCbo()
    Set rsColor = New ADODB.Recordset
    rsColor.Open "Select COLOR_DESC from ALL_Color", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsColor.EOF And Not rsColor.BOF Then
        rsColor.MoveFirst
        cboCOLOR.Clear
        Do While Not rsColor.EOF
            cboCOLOR.AddItem Null2String(rsColor!color_desc)
            rsColor.MoveNext
        Loop
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
    Frame1.Enabled = False
    Picture2.Enabled = False
    txtINSTYPE.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim vtxtCusCde                                      As String
    Dim VTXTNiym                                        As String
    Dim VTXTPlateNo                                     As String
    Dim VtxtProdNo                                      As String
    Dim VtxtVCond_No                                    As String
    Dim VcboModel                                       As String
    Dim Vcbocolor                                       As String
    Dim VtxtEngine                                      As String
    Dim VtxtSerial                                      As String
    Dim VtxtD_Sold                                      As String
    Dim VtxtWar_Cert                                    As String
    Dim VtxtTin_Number                                  As String
    Dim VtxtDel_Date                                    As String
    Dim vSellingDealer                                  As String
    Dim vEndUser                                        As String
    Dim vINSTYPE                                        As String
    Dim vINSCOMP                                        As String
    Dim VINSDATE                                        As String
    Dim vFINTYPE                                        As String
    Dim vFINCOMP                                        As String
    Dim vFINDATE                                        As String
    Dim RSTMP                                           As New ADODB.Recordset
    
    Set RSTMP = gconDMIS.Execute("SELECT PLATE_NO, ID FROM CSMS_CUSVEH WHERE PLATE_NO = " & N2Str2Null(txtPLATENO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Not (labID.Caption = RSTMP!ID) Then
            MsgBox " This Plate no Already Exist, Editing this plate no can cause duplicate record", vbExclamation, "Error"
            txtPLATENO.SetFocus
            Exit Sub
        End If
    End If
    
    Set RSTMP = gconDMIS.Execute("SELECT VIN, ID FROM CSMS_CUSVEH WHERE VIN = " & N2Str2Null(txtSerial) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Not (labID.Caption = RSTMP!ID) Then
            MsgBox "Vin no Already exist,  Editing this Vin no can cause duplicate record", vbExclamation, "Error"
            txtSerial.SetFocus
            Exit Sub
        End If
    End If
    Set RSTMP = Nothing
    
    If txtMake.Text = "" Then
        ShowIsRequiredMsg ("Make Cannot be Blank")
        On Error Resume Next
        cmdSelect.SetFocus
        Exit Sub
    End If
    
    vtxtCusCde = N2Str2Null(xLOCAL_CUSCDE)
    VTXTNiym = N2Str2Null(xLOCAL_ACCTNAME)
    VTXTPlateNo = N2Str2Null(txtPLATENO.Text)
    
    VtxtProdNo = N2Str2Null(txtProdNo.Text)
    VcboModel = N2Str2Null(txtModel.Text)
    Vcbocolor = N2Str2Null(SetColor(cboCOLOR.Text))
    VtxtVCond_No = N2Str2Null(txtVCond_No.Text)
    VtxtEngine = N2Str2Null(txtENGINE.Text)
    VtxtSerial = N2Str2Null(txtSerial.Text)
    vSellingDealer = N2Str2Null(SetSellingDealer(cboSelling_Dealer, 2))
    vEndUser = N2Str2Null(SetEndUser(cboEndUser.Text, 2))
    If IsDate(txtD_Sold.Text) = True Then
        VtxtD_Sold = N2Str2Null(Format(txtD_Sold.Text, "short date"))
    Else
        VtxtD_Sold = "NULL"
    End If
    VtxtWar_Cert = N2Str2Null(txtWar_Cert.Text)
    VtxtTin_Number = N2Str2Null(txtTin_Number.Text)
    If IsDate(txtDel_Date.Text) = True Then
        VtxtDel_Date = N2Str2Null(Format(txtDel_Date.Text, "Short date"))
    Else
        VtxtDel_Date = "NULL"
    End If
    vINSTYPE = N2Str2Null(txtINSTYPE)
    vINSCOMP = N2Str2Null(txtINSCOMP)
    VINSDATE = N2Date2Null(dtpINS.Value)
    vFINTYPE = N2Str2Null(txtINSTYPE)
    vFINCOMP = N2Str2Null(txtFINCOMP)
    vFINDATE = N2Date2Null(dtpFIN.Value)
    
    If AddorEdit = "ADD" Then
        If IsNull(txtProdNo.Text) = False Then
            Dim rsCusVehDup                            As New ADODB.Recordset
            rsCusVehDup.Open "select prodno from CSMS_CusVeh where prodno = '" & txtProdNo.Text & "'", gconDMIS
            If Not rsCusVehDup.EOF And Not rsCusVeh.BOF Then
                MsgSpeechBox "Product Number Already Exist"
                Exit Sub
            End If
        End If
        
        SQL_STATEMENT = "insert into CSMS_CusVeh " & _
            "(cuscde, niym, YER, Make, Description, plate_no, prodno, vcond_no, model, clrcde, engine, serial, tin_number, d_sold, war_cert, del_date, Selling_Dealer, EndUser)" & _
            " values (" & vtxtCusCde & _
            ", " & VTXTNiym & _
            ", " & N2Str2Null(txtYear.Text) & _
            ", " & N2Str2Null(txtMake.Text) & _
            ", " & N2Str2Null(txtDescription.Text) & _
            ", " & VTXTPlateNo & _
            ", " & VtxtProdNo & _
            ", " & VtxtVCond_No & _
            ", " & VcboModel & _
            ", " & Vcbocolor & _
            ", " & VtxtEngine & _
            ", " & VtxtSerial & _
            ", " & VtxtTin_Number & _
            ", " & VtxtD_Sold & _
            ", " & VtxtWar_Cert & _
            ", " & VtxtDel_Date & _
            "," & vSellingDealer & _
            "," & vEndUser & ")"
        gconDMIS.Execute SQL_STATEMENT

        Set rsCustomer = New ADODB.Recordset
        rsCustomer.Open "select id,cuscde from ALL_Customer where cuscde = " & vtxtCusCde, gconDMIS
        If Not rsCustomer.EOF And Not rsCustomer.BOF Then
            gconDMIS.Execute "update ALL_Customer set plateno = " & VTXTPlateNo & " where id = " & rsCustomer!ID
        End If
    Else
        Set RSTMP = gconDMIS.Execute("SELECT PLATE_NO FROM CSMS_REPOR WHERE PLATE_NO = " & N2Str2Null(xOLD_PLATENO) & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            If MsgBox("This Vehicle has an History Record in Repair Order/Estimate, updating vehicle information will also update the Information in the History File (e.g. Plate, Vin no...), Proceed", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
            
            gconDMIS.Execute ("UPDATE CSMS_REPOR SET PLATE_NO = " & VTXTPlateNo & _
                ", MODEL = " & VcboModel & _
                ", VIN = " & VtxtSerial & _
                " WHERE PLATE_NO = " & N2Str2Null(xOLD_PLATENO) & "")
                
            gconDMIS.Execute ("UPDATE CSMS_REPAIRORDER SET PLATE_NO = " & VTXTPlateNo & _
                ", MODEL = " & VcboModel & _
                " WHERE PLATE_NO = " & N2Str2Null(xOLD_PLATENO) & "")
                
            gconDMIS.Execute ("UPDATE CSMS_APPOINTMENT SET PLATE_NO = " & VTXTPlateNo & _
                ", MODEL = " & VcboModel & _
                ", MAKE = " & N2Str2Null(txtMake) & _
                " WHERE PLATE_NO = " & N2Str2Null(xOLD_PLATENO) & "")
                
            gconDMIS.Execute ("UPDATE CSMS_ESTHD SET PLATE_NO = " & VTXTPlateNo & _
                ", MODEL = " & VcboModel & _
                ", VIN = " & VtxtSerial & _
                " WHERE PLATE_NO = " & N2Str2Null(xOLD_PLATENO) & "")
        End If
        
        SQL_STATEMENT = "update CSMS_CusVeh set" & _
            " PLATE_NO = " & VTXTPlateNo & _
            ", Yer = " & N2Str2Null(txtYear.Text) & _
            ", Make = " & N2Str2Null(txtMake.Text) & _
            ", VCond_no = " & VtxtVCond_No & _
            ", prodno = " & VtxtProdNo & _
            ", model = " & VcboModel & _
            ", clrcde = " & Vcbocolor & _
            ", engine = " & VtxtEngine & _
            ", VIN = " & VtxtSerial & _
            ", serial = " & VtxtSerial & _
            ", tin_number = " & VtxtTin_Number & _
            ", d_sold = " & VtxtD_Sold & _
            ", war_cert = " & VtxtWar_Cert & _
            ", Description = " & N2Str2Null(txtDescription.Text) & _
            ", Selling_Dealer = " & vSellingDealer & _
            ", EndUser = " & vEndUser & _
            ", del_date = " & VtxtDel_Date & _
            ", INS_TYPE = " & vINSTYPE & _
            ", INS_COMP = " & vINSCOMP & _
            ", INS_EXP_DATE = " & VINSDATE & _
            ", FIN_TYPE = " & vFINTYPE & _
            ", FIN_COMP = " & vFINCOMP & _
            ", FIN_EXP_DATE = " & vFINDATE & _
            " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("E", "CUSTOMER VEHICLE", SQL_STATEMENT, labID, "", "COND NO: " & txtVCond_No, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        Call ShowSuccessFullyUpdated
    End If

    RaiseEvent SaveChanges(txtPLATENO, txtWar_Cert, txtMake, txtModel, txtSerial, txtDescription, xFromForm)
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdSelect_Click()
    'frmCSMSROYrMkMlEgn.Show 1

    Set frmSelectMakeMode = New frmCSMSYrMkMlEgn
    frmSelectMakeMode.Show 1
End Sub

Private Sub frmSelectMakeMode_SelectedDetails(XYEAR As String, xMake As String, xMODEL As String, xENGINE As String, xModelDescription As String)
    txtYear = XYEAR
    txtMake = xMake
    txtModel = xMODEL
    txtENGINE = xENGINE
    txtDescription = xModelDescription
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    
    Call FillCboSellingDealer
    Call FillCboEndUser
    Call initMemvars
    
    'Set rsCusVeh = New ADODB.Recordset
    'rsCusVeh.Open "select * from CSMS_CusVeh where plate_no = '" & frmCSMSDataEntry.txtPlate_No.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Call StoreMemVars
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmCSMSDataEntry.Enabled = True
    'Set frmCSMSROCusveh = Nothing
End Sub

Private Sub txtD_Sold_LostFocus()
    If txtD_Sold.Text <> "" Then txtD_Sold.Text = Format(txtD_Sold.Text, "Short Date")
End Sub

Private Sub txtDel_Date_LostFocus()
    If txtDel_Date.Text <> "" Then txtDel_Date.Text = Format(txtDel_Date.Text, "Short Date")
End Sub

