VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMS_PartsInquiry 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PARTS PRICE LOOKUP INQUIRY"
   ClientHeight    =   6930
   ClientLeft      =   315
   ClientTop       =   330
   ClientWidth     =   10845
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "PartsInquiry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10845
   Begin VB.PictureBox picDESC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6885
      Left            =   30
      ScaleHeight     =   6825
      ScaleWidth      =   10755
      TabIndex        =   0
      Top             =   30
      Width           =   10815
      Begin VB.CommandButton cmdInquire 
         Caption         =   "Inquire by Part Description"
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
         Left            =   8130
         TabIndex        =   140
         Top             =   60
         Width           =   2595
      End
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
         Height          =   5655
         Left            =   30
         TabIndex        =   5
         Top             =   390
         Width           =   10695
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   20
            ToolTipText     =   "Input Valid Part Number"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   90
            TabIndex        =   19
            ToolTipText     =   "Input Valid Part Number"
            Top             =   930
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   2
            Left            =   90
            TabIndex        =   18
            ToolTipText     =   "Input Valid Part Number"
            Top             =   1260
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   3
            Left            =   90
            TabIndex        =   17
            ToolTipText     =   "Input Valid Part Number"
            Top             =   1590
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   90
            TabIndex        =   16
            ToolTipText     =   "Input Valid Part Number"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   90
            TabIndex        =   15
            ToolTipText     =   "Input Valid Part Number"
            Top             =   2250
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   6
            Left            =   90
            TabIndex        =   14
            ToolTipText     =   "Input Valid Part Number"
            Top             =   2580
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   7
            Left            =   90
            TabIndex        =   13
            ToolTipText     =   "Input Valid Part Number"
            Top             =   2910
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   8
            Left            =   90
            TabIndex        =   12
            ToolTipText     =   "Input Valid Part Number"
            Top             =   3240
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   9
            Left            =   90
            TabIndex        =   11
            ToolTipText     =   "Input Valid Part Number"
            Top             =   3570
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   10
            Left            =   90
            TabIndex        =   10
            ToolTipText     =   "Input Valid Part Number"
            Top             =   3900
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   11
            Left            =   90
            TabIndex        =   9
            ToolTipText     =   "Input Valid Part Number"
            Top             =   4230
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   12
            Left            =   90
            TabIndex        =   8
            ToolTipText     =   "Input Valid Part Number"
            Top             =   4560
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   13
            Left            =   90
            TabIndex        =   7
            ToolTipText     =   "Input Valid Part Number"
            Top             =   4890
            Width           =   1575
         End
         Begin VB.TextBox txtPartNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   14
            Left            =   90
            TabIndex        =   6
            ToolTipText     =   "Input Valid Part Number"
            Top             =   5220
            Width           =   1575
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   133
            Top             =   630
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   132
            Top             =   630
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   5280
            TabIndex        =   131
            Top             =   630
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   6810
            TabIndex        =   130
            Top             =   630
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   7800
            TabIndex        =   129
            Top             =   630
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   128
            Top             =   990
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   127
            Top             =   990
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   5280
            TabIndex        =   126
            Top             =   990
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   6810
            TabIndex        =   125
            Top             =   990
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   7800
            TabIndex        =   124
            Top             =   990
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   123
            Top             =   1320
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   122
            Top             =   1320
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   5280
            TabIndex        =   121
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   6810
            TabIndex        =   120
            Top             =   1320
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   7800
            TabIndex        =   119
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   118
            Top             =   1650
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   117
            Top             =   1650
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   5280
            TabIndex        =   116
            Top             =   1650
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   6810
            TabIndex        =   115
            Top             =   1650
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   7800
            TabIndex        =   114
            Top             =   1650
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   1680
            TabIndex        =   113
            Top             =   1980
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   112
            Top             =   1980
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   5280
            TabIndex        =   111
            Top             =   1980
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   6810
            TabIndex        =   110
            Top             =   1980
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   7800
            TabIndex        =   109
            Top             =   1980
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   1680
            TabIndex        =   108
            Top             =   2310
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   3600
            TabIndex        =   107
            Top             =   2310
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   5280
            TabIndex        =   106
            Top             =   2310
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   6810
            TabIndex        =   105
            Top             =   2310
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   7800
            TabIndex        =   104
            Top             =   2310
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   1680
            TabIndex        =   103
            Top             =   2640
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   3600
            TabIndex        =   102
            Top             =   2640
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   5280
            TabIndex        =   101
            Top             =   2640
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   6810
            TabIndex        =   100
            Top             =   2640
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   7800
            TabIndex        =   99
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   1680
            TabIndex        =   98
            Top             =   2970
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   3600
            TabIndex        =   97
            Top             =   2970
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   5280
            TabIndex        =   96
            Top             =   2970
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   6810
            TabIndex        =   95
            Top             =   2970
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   7800
            TabIndex        =   94
            Top             =   2970
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   1680
            TabIndex        =   93
            Top             =   3300
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   3600
            TabIndex        =   92
            Top             =   3300
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   5280
            TabIndex        =   91
            Top             =   3300
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   6810
            TabIndex        =   90
            Top             =   3300
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   7800
            TabIndex        =   89
            Top             =   3300
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   1680
            TabIndex        =   88
            Top             =   3630
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   3600
            TabIndex        =   87
            Top             =   3630
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   5280
            TabIndex        =   86
            Top             =   3630
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   6810
            TabIndex        =   85
            Top             =   3630
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   7800
            TabIndex        =   84
            Top             =   3630
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   1680
            TabIndex        =   83
            Top             =   3960
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   3600
            TabIndex        =   82
            Top             =   3960
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   5280
            TabIndex        =   81
            Top             =   3960
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   6810
            TabIndex        =   80
            Top             =   3960
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   7800
            TabIndex        =   79
            Top             =   3960
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   1680
            TabIndex        =   78
            Top             =   4290
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   3600
            TabIndex        =   77
            Top             =   4290
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   5280
            TabIndex        =   76
            Top             =   4290
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   6810
            TabIndex        =   75
            Top             =   4290
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   7800
            TabIndex        =   74
            Top             =   4290
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   1680
            TabIndex        =   73
            Top             =   4620
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   3600
            TabIndex        =   72
            Top             =   4620
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   5280
            TabIndex        =   71
            Top             =   4620
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   6810
            TabIndex        =   70
            Top             =   4620
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   7800
            TabIndex        =   69
            Top             =   4620
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   1680
            TabIndex        =   68
            Top             =   4950
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   3600
            TabIndex        =   67
            Top             =   4950
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   5280
            TabIndex        =   66
            Top             =   4950
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   6810
            TabIndex        =   65
            Top             =   4950
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   7800
            TabIndex        =   64
            Top             =   4950
            Width           =   495
         End
         Begin VB.Label txtDescrip 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   1680
            TabIndex        =   63
            Top             =   5280
            Width           =   2025
         End
         Begin VB.Label txtDNPP 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   3600
            TabIndex        =   62
            Top             =   5280
            Width           =   1755
         End
         Begin VB.Label txtSRP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   5280
            TabIndex        =   61
            Top             =   5280
            Width           =   1395
         End
         Begin VB.Label txtModel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   6810
            TabIndex        =   60
            Top             =   5280
            Width           =   945
         End
         Begin VB.Label txtICC 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   7800
            TabIndex        =   59
            Top             =   5280
            Width           =   495
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   1680
            TabIndex        =   58
            Top             =   240
            Width           =   1905
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SUPERCESSION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   3600
            TabIndex        =   57
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   5400
            TabIndex        =   56
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MODEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   6810
            TabIndex        =   55
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ICC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   7800
            TabIndex        =   54
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label6 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "PART NUMBER"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   90
            TabIndex        =   53
            Top             =   240
            Width           =   1545
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000000&
            X1              =   30
            X2              =   10590
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   8220
            TabIndex        =   52
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   8370
            TabIndex        =   51
            Top             =   5280
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   8370
            TabIndex        =   50
            Top             =   4950
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   8370
            TabIndex        =   49
            Top             =   4620
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   8370
            TabIndex        =   48
            Top             =   4290
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   8370
            TabIndex        =   47
            Top             =   3960
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   8370
            TabIndex        =   46
            Top             =   3630
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   8370
            TabIndex        =   45
            Top             =   3300
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   8370
            TabIndex        =   44
            Top             =   2970
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   8370
            TabIndex        =   43
            Top             =   2640
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   8370
            TabIndex        =   42
            Top             =   2310
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   8370
            TabIndex        =   41
            Top             =   1980
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   8370
            TabIndex        =   40
            Top             =   1650
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   8370
            TabIndex        =   39
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   8370
            TabIndex        =   38
            Top             =   990
            Width           =   795
         End
         Begin VB.Label txtSTOCK 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   8370
            TabIndex        =   37
            Top             =   630
            Width           =   795
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   9240
            TabIndex        =   36
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   9240
            TabIndex        =   35
            Top             =   5280
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   9240
            TabIndex        =   34
            Top             =   4950
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   9240
            TabIndex        =   33
            Top             =   4620
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   9240
            TabIndex        =   32
            Top             =   4290
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   9240
            TabIndex        =   31
            Top             =   3960
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   9240
            TabIndex        =   30
            Top             =   3630
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   9240
            TabIndex        =   29
            Top             =   3300
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   9240
            TabIndex        =   28
            Top             =   2970
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   9240
            TabIndex        =   27
            Top             =   2640
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   9240
            TabIndex        =   26
            Top             =   2310
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   9240
            TabIndex        =   25
            Top             =   1980
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   9240
            TabIndex        =   24
            Top             =   1650
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   9240
            TabIndex        =   23
            Top             =   1320
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   9240
            TabIndex        =   22
            Top             =   990
            Width           =   1365
         End
         Begin VB.Label txtLOCATION 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "LOCATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   9240
            TabIndex        =   21
            Top             =   630
            Width           =   1365
         End
      End
      Begin VB.PictureBox Picture1 
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
         Height          =   375
         Left            =   30
         ScaleHeight     =   375
         ScaleWidth      =   7425
         TabIndex        =   1
         Top             =   60
         Width           =   7425
         Begin VB.OptionButton Option1 
            Caption         =   "Stock Option"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   30
            TabIndex        =   4
            Top             =   60
            Value           =   -1  'True
            Width           =   1845
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Dealer Master File Only"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1680
            TabIndex        =   3
            Top             =   60
            Visible         =   0   'False
            Width           =   2595
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Distributor Master File Only"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4560
            TabIndex        =   2
            Top             =   60
            Visible         =   0   'False
            Width           =   3105
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
         Left            =   570
         TabIndex        =   139
         Top             =   540
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label labid 
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
         Left            =   240
         TabIndex        =   138
         Top             =   690
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label labNote 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "NOTE:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   15
         Left            =   60
         TabIndex        =   137
         Top             =   6120
         Width           =   675
      End
      Begin VB.Label labNote 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   $"PartsInquiry.frx":030A
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   705
         Index           =   0
         Left            =   750
         TabIndex        =   136
         Top             =   6120
         Width           =   9885
      End
      Begin VB.Label labNote 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "BLUE COLOR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   5910
         TabIndex        =   135
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Label labNote 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "RED COLOR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   7755
         TabIndex        =   134
         Top             =   6315
         Width           =   1335
      End
   End
   Begin VB.PictureBox picINQ 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6825
      Left            =   30
      ScaleHeight     =   6795
      ScaleWidth      =   10725
      TabIndex        =   141
      Top             =   30
      Visible         =   0   'False
      Width           =   10755
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   4275
         Left            =   5520
         ScaleHeight     =   4275
         ScaleWidth      =   5115
         TabIndex        =   148
         Top             =   360
         Width           =   5115
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   345
            Left            =   30
            TabIndex        =   165
            Top             =   30
            Width           =   5055
            _Version        =   655364
            _ExtentX        =   8916
            _ExtentY        =   609
            _StockProps     =   14
            Caption         =   "Inventory Information"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            VisualTheme     =   3
         End
         Begin VB.Label lblCAP 
            Appearance      =   0  'Flat
            BackColor       =   &H00D2BDB6&
            Caption         =   "Stock no"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   0
            Left            =   30
            TabIndex        =   164
            Top             =   390
            Width           =   1860
         End
         Begin VB.Label lblCAP 
            Appearance      =   0  'Flat
            BackColor       =   &H00D2BDB6&
            Caption         =   "Stock Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   1
            Left            =   30
            TabIndex        =   163
            Top             =   690
            Width           =   1860
         End
         Begin VB.Label lblCAP 
            Appearance      =   0  'Flat
            BackColor       =   &H00D2BDB6&
            Caption         =   "New Part No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   2
            Left            =   30
            TabIndex        =   162
            Top             =   990
            Width           =   1860
         End
         Begin VB.Label lblCAP 
            Appearance      =   0  'Flat
            BackColor       =   &H00D2BDB6&
            Caption         =   "SRP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   3
            Left            =   30
            TabIndex        =   161
            Top             =   1290
            Width           =   1860
         End
         Begin VB.Label lblCAP 
            Appearance      =   0  'Flat
            BackColor       =   &H00D2BDB6&
            Caption         =   "Model Code"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   735
            Index           =   4
            Left            =   30
            TabIndex        =   160
            Top             =   1590
            Width           =   1860
         End
         Begin VB.Label lblCAP 
            Appearance      =   0  'Flat
            BackColor       =   &H00D2BDB6&
            Caption         =   "Inv. Class"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   5
            Left            =   30
            TabIndex        =   159
            Top             =   2340
            Width           =   1860
         End
         Begin VB.Label lblCAP 
            Appearance      =   0  'Flat
            BackColor       =   &H00D2BDB6&
            Caption         =   "Stock"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   6
            Left            =   30
            TabIndex        =   158
            Top             =   2640
            Width           =   1860
         End
         Begin VB.Label lblCAP 
            Appearance      =   0  'Flat
            BackColor       =   &H00D2BDB6&
            Caption         =   "Location"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   1275
            Index           =   7
            Left            =   30
            TabIndex        =   157
            Top             =   2940
            Width           =   1860
         End
         Begin VB.Label lblres 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   156
            Top             =   390
            Width           =   3165
         End
         Begin VB.Label lblres 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   1920
            TabIndex        =   155
            Top             =   690
            Width           =   3165
         End
         Begin VB.Label lblres 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   1920
            TabIndex        =   154
            Top             =   990
            Width           =   3165
         End
         Begin VB.Label lblres 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   3
            Left            =   1920
            TabIndex        =   153
            Top             =   1290
            Width           =   3165
         End
         Begin VB.Label lblres 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Index           =   4
            Left            =   1920
            TabIndex        =   152
            Top             =   1590
            Width           =   3165
         End
         Begin VB.Label lblres 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   5
            Left            =   1920
            TabIndex        =   151
            Top             =   2340
            Width           =   3165
         End
         Begin VB.Label lblres 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   6
            Left            =   1920
            TabIndex        =   150
            Top             =   2640
            Width           =   3165
         End
         Begin VB.Label lblres 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1275
            Index           =   7
            Left            =   1920
            TabIndex        =   149
            Top             =   2940
            Width           =   3165
         End
      End
      Begin VB.ComboBox cboModel 
         Height          =   330
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   146
         Top             =   810
         Width           =   3855
      End
      Begin VB.CheckBox chkModel 
         Caption         =   "By Model"
         Height          =   225
         Left            =   150
         TabIndex        =   145
         Top             =   840
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "X"
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
         Left            =   10380
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   0
         Width           =   345
      End
      Begin MSComctlLib.ListView lsvLIST 
         Height          =   5535
         Left            =   90
         TabIndex        =   143
         Top             =   1170
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   9763
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
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PART NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PART DESCRIPTION"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox txtSeach 
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
         Left            =   90
         TabIndex        =   142
         Top             =   360
         Width           =   5385
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   345
         Left            =   -30
         TabIndex        =   147
         Top             =   -30
         Width           =   10785
         _Version        =   655364
         _ExtentX        =   19024
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "   INQUIRE PART AVAILABILITY"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
      End
   End
End
Attribute VB_Name = "frmCSMS_PartsInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub InquireIT()
    Dim rsDNPP                                         As ADODB.Recordset
    Dim rsPartMas                                      As ADODB.Recordset
    Dim SpeakSTOCKNUMBER                               As String
    Dim I                                              As Long
    Dim kim                                            As Long
    For I = 0 To 14
        If txtPartNo(I).Text <> "" Then
            LogAudit "V", "PARTS INQURIY ", "FOR STOCKNO" & txtSTOCK(I)
            txtDescrip(I).BorderStyle = 1
            txtDNPP(I).BorderStyle = 1
            txtSRP(I).BorderStyle = 1
            txtModel(I).BorderStyle = 1
            txtICC(I).BorderStyle = 1
            txtSTOCK(I).BorderStyle = 1
            txtLOCATION(I).BorderStyle = 1

            If Option1.Value = True Then
                Set rsPartMas = New ADODB.Recordset
                rsPartMas.Open "select STOCKNO,STOCKDESC,dnp,newno,srp,modelcode,location,invclass,subinvclas,onhand from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(txtPartNo(I).Text) & " and onhand > 0", gconDMIS
                If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                    txtDescrip(I).ForeColor = vbBlue: txtDNPP(I).ForeColor = vbBlue: txtSRP(I).ForeColor = vbBlue: txtModel(I).ForeColor = vbBlue
                    txtICC(I).ForeColor = vbBlue: txtSTOCK(I).ForeColor = vbBlue: txtLOCATION(I).ForeColor = vbBlue
                    txtDescrip(I).Caption = Null2String(rsPartMas!STOCKDESC): txtDNPP(I).Caption = Null2String(rsPartMas!newno)
                    txtSRP(I).Caption = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP)): txtModel(I).Caption = Null2String(rsPartMas!MODELCODE)
                    txtICC(I).Caption = Null2String(rsPartMas!InvClass) & Null2String(rsPartMas!SubInvClas)
                    txtSTOCK(I).Caption = N2Str2Zero(rsPartMas!ONHAND):
                    'BTT - 05292007
                    If N2Str2Zero(rsPartMas!ONHAND) > 1 Then
                        txtSTOCK(I).Caption = "Yes"
                    Else
                        txtSTOCK(I).Caption = "N/A"
                    End If
                    txtLOCATION(I).Caption = Null2String(rsPartMas!Location)
                    SpeakSTOCKNUMBER = "": For kim = 1 To Len(txtPartNo(I).Text): SpeakSTOCKNUMBER = SpeakSTOCKNUMBER & Mid(txtPartNo(I).Text, kim, 1) & " ": Next

                Else
                    Set rsDNPP = New ADODB.Recordset
                    rsDNPP.Open "select * from PMIS_DNPP where NEWPARTNO=" & N2Str2Null(txtPartNo(I).Text), gconDMIS
                    If Not rsDNPP.EOF And Not rsDNPP.BOF Then
                        txtDescrip(I).ForeColor = vbRed: txtDNPP(I).ForeColor = vbRed: txtSRP(I).ForeColor = vbRed
                        txtModel(I).ForeColor = vbRed: txtICC(I).ForeColor = vbRed: txtSTOCK(I).ForeColor = vbRed: txtLOCATION(I).ForeColor = vbRed
                        txtDescrip(I).Caption = Null2String(rsDNPP!DESCRIPTIO): txtDNPP(I).Caption = Null2String(rsDNPP!STOCKNUMBER)
                        txtSRP(I).Caption = ToDoubleNumber(N2Str2Zero(rsDNPP!SRP)): txtModel(I).Caption = Null2String(rsDNPP!MODEL)
                        txtICC(I).Caption = Null2String(rsDNPP!icc): txtSTOCK(I).Caption = "": txtLOCATION(I).Caption = ""

                    Else
                        Set rsDNPP = New ADODB.Recordset
                        rsDNPP.Open "select * from PMIS_DNPP where PARTNUMBER=" & N2Str2Null(txtPartNo(I).Text), gconDMIS
                        If Not rsDNPP.EOF And Not rsDNPP.BOF Then
                            txtDescrip(I).ForeColor = vbRed: txtDNPP(I).ForeColor = vbRed: txtSRP(I).ForeColor = vbRed
                            txtModel(I).ForeColor = vbRed: txtICC(I).ForeColor = vbRed: txtSTOCK(I).ForeColor = vbRed: txtLOCATION(I).ForeColor = vbRed
                            txtDescrip(I).Caption = Null2String(rsDNPP!DESCRIPTIO): txtDNPP(I).Caption = Null2String(rsDNPP!NewPARTNO)
                            txtSRP(I).Caption = ToDoubleNumber(N2Str2Zero(rsDNPP!SRP)): txtModel(I).Caption = Null2String(rsDNPP!MODEL)
                            txtICC(I).Caption = Null2String(rsDNPP!icc): txtSTOCK(I).Caption = "": txtLOCATION(I).Caption = ""
                        Else
                            txtDescrip(I).Caption = "Not in Master"
                            txtDNPP(I).Caption = "": txtSRP(I).Caption = "": txtModel(I).Caption = ""
                            txtICC(I).Caption = "": txtSTOCK(I).Caption = "": txtLOCATION(I).Caption = ""
                        End If
                    End If
                End If
            End If

            If Option2.Value = True Then
                Set rsPartMas = New ADODB.Recordset
                rsPartMas.Open "select STOCKNO,STOCKDESC,dnp,newno,srp,modelcode,location,invclass,subinvclas,onhand from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(txtPartNo(I).Text) & " and onhand > 0", gconDMIS
                If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                    txtDescrip(I).ForeColor = vbBlue: txtDNPP(I).ForeColor = vbBlue: txtSRP(I).ForeColor = vbBlue: txtModel(I).ForeColor = vbBlue
                    txtICC(I).ForeColor = vbBlue: txtSTOCK(I).ForeColor = vbBlue: txtLOCATION(I).ForeColor = vbBlue
                    txtDescrip(I).Caption = Null2String(rsPartMas!STOCKDESC): txtDNPP(I).Caption = Null2String(rsPartMas!newno)
                    txtSRP(I).Caption = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP)): txtModel(I).Caption = Null2String(rsPartMas!MODELCODE)
                    txtICC(I).Caption = Null2String(rsPartMas!InvClass) & Null2String(rsPartMas!SubInvClas)
                    txtSTOCK(I).Caption = N2Str2Zero(rsPartMas!ONHAND): txtLOCATION(I).Caption = Null2String(rsPartMas!Location)
                    SpeakSTOCKNUMBER = "": For kim = 1 To Len(txtPartNo(I).Text): SpeakSTOCKNUMBER = SpeakSTOCKNUMBER & Mid(txtPartNo(I).Text, kim, 1) & " ": Next
                Else
                    txtDescrip(I).Caption = "Not in Master"
                    txtDNPP(I).Caption = "": txtSRP(I).Caption = "": txtModel(I).Caption = ""
                    txtICC(I).Caption = "": txtSTOCK(I).Caption = "": txtLOCATION(I).Caption = ""
                End If
            End If
            If Option3.Value = True Then
                Set rsDNPP = New ADODB.Recordset
                rsDNPP.Open "select * from PMIS_DNPP where PARTNUMBER=" & N2Str2Null(txtPartNo(I).Text), gconDMIS
                If Not rsDNPP.EOF And Not rsDNPP.BOF Then
                    txtDescrip(I).ForeColor = vbRed: txtDNPP(I).ForeColor = vbRed: txtSRP(I).ForeColor = vbRed
                    txtModel(I).ForeColor = vbRed: txtICC(I).ForeColor = vbRed: txtSTOCK(I).ForeColor = vbRed: txtLOCATION(I).ForeColor = vbRed
                    txtDescrip(I).Caption = Null2String(rsDNPP!DESCRIPTIO): txtDNPP(I).Caption = Null2String(rsDNPP!newSTOCKNO)
                    txtSRP(I).Caption = ToDoubleNumber(N2Str2Zero(rsDNPP!SRP)): txtModel(I).Caption = Null2String(rsDNPP!MODEL)
                    txtICC(I).Caption = Null2String(rsDNPP!icc): txtSTOCK(I).Caption = "": txtLOCATION(I).Caption = ""
                Else
                    txtDescrip(I).Caption = "Not in Master"
                    txtDNPP(I).Caption = "": txtSRP(I).Caption = "": txtModel(I).Caption = ""
                    txtICC(I).Caption = "": txtSTOCK(I).Caption = "": txtLOCATION(I).Caption = ""
                End If
            End If
            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("I", "PARTS INQUIRY", "", "", "", "PART NO: " & txtPartNo(I).Text, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        Else
            txtDescrip(I).Caption = ""
            txtDNPP(I).Caption = ""
            txtSRP(I).Caption = ""
            txtModel(I).Caption = ""
            txtICC(I).Caption = ""
            txtSTOCK(I).Caption = ""
            txtLOCATION(I).Caption = ""
            txtDescrip(I).BorderStyle = 0
            txtDNPP(I).BorderStyle = 0
            txtSRP(I).BorderStyle = 0
            txtModel(I).BorderStyle = 0
            txtICC(I).BorderStyle = 0
            txtSTOCK(I).BorderStyle = 0
            txtLOCATION(I).BorderStyle = 0
        End If
    Next


End Sub

Sub initMemvars()
    Dim k                                              As Integer
    For k = 0 To 14
        txtPartNo(k).Text = ""
        txtDescrip(k).Caption = ""
        txtDNPP(k).Caption = ""
        txtSRP(k).Caption = ""
        txtModel(k).Caption = ""
        txtICC(k).Caption = ""
        txtSTOCK(k).Caption = ""
        txtLOCATION(k).Caption = ""
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInquire_Click()
    picDESC.Enabled = False
    picINQ.Visible = True
    picINQ.ZOrder 0
    txtSeach.Text = ""
    txtSeach.SetFocus
End Sub

Private Sub Command1_Click()
    picDESC.Enabled = True
    picINQ.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        InquireIT
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PARTS INQUIRY)"
            Call frmALL_AuditInquiry.DisplayHistory("", "PARTS INQUIRY", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Dim k                                              As Integer
    initMemvars
    Call FillModel
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCSMS_PartsInquiry = Nothing
    Unload Me
End Sub

Private Sub lblEmails_Click()

End Sub

Private Sub lsvLIST_ItemCheck(ByVal ITEM As MSComctlLib.ListItem)
'    Dim RSTMP As New ADODB.Recordset
'    Set RSTMP = gconDMIS.Execute("SELECT TOP 100 *,STOCKNO, STOCKDESC,DNP,NEWNO,SRP,MODELCODE,LOCATION,INVCLASS,subinvclas,onhand FROM PMIS_StockMas WHERE ID = " & Item.ListSubItems(2) & "")
'    If Not (RSTMP.BOF And RSTMP.EOF) Then
'        lblres(0).Caption = Null2String(RSTMP!STOCKNO)
'        lblres(1).Caption = Null2String(RSTMP!STOCKDESC)
'        lblres(2).Caption = NumericVal(RSTMP!NEWNO)
'        lblres(3).Caption = NumericVal(RSTMP!SRP)
'        lblres(4).Caption = Null2String(RSTMP!MODELCODE)
'        lblres(5).Caption = Null2String(RSTMP!INVCLASS)
'        If NumericVal(RSTMP!ONHAND) < 1 Then
'            lblres(6).Caption = "N"
'        Else
'            lblres(6).Caption = "Y"
'        End If
'        lblres(7).Caption = Null2String(RSTMP!Location)
'    End If
'    Set RSTMP = Nothing
End Sub

Sub FillModel()
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT CODE, DESCRIPT FROM CSMS_MODELS ORDER BY CODE")
    cboModel.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboModel.AddItem Null2String(RSTMP!Code) & " - " & Null2String(RSTMP!DESCRIPT)
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Private Sub lsvLIST_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT TOP 100 *,STOCKNO, STOCKDESC,DNP,NEWNO,SRP,MODELCODE,LOCATION,INVCLASS,subinvclas,onhand FROM PMIS_StockMas WHERE ID = " & ITEM.ListSubItems(2) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        lblRES(0).Caption = Null2String(RSTMP!STOCKNO)
        lblRES(1).Caption = Null2String(RSTMP!STOCKDESC)
        lblRES(2).Caption = Null2String(RSTMP!newno)
        lblRES(3).Caption = NumericVal(RSTMP!SRP)
        lblRES(4).Caption = Null2String(RSTMP!MODELCODE)
        lblRES(5).Caption = Null2String(RSTMP!InvClass)
        If NumericVal(RSTMP!ONHAND) < 1 Then
            lblRES(6).Caption = "N"
        Else
            lblRES(6).Caption = "Y"
        End If
        lblRES(7).Caption = Null2String(RSTMP!Location)
    End If
    Set RSTMP = Nothing
End Sub

Private Sub txtPartNo_KeyPress(INDEX As Integer, KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSeach_Change()
    Call FillSearchGrid(txtSeach)
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSTMP As New ADODB.Recordset
    Dim ITEM As ListItem
    
    XXX = Replace(XXX, "'", "")
    If XXX = "" Then
        Set RSTMP = gconDMIS.Execute("SELECT TOP 100 *,STOCKNO, STOCKDESC,DNP,NEWNO,SRP,MODELCODE,LOCATION,INVCLASS,subinvclas,onhand FROM PMIS_StockMas")
    Else
        If chkModel.Value = 1 Then
            Set RSTMP = gconDMIS.Execute("SELECT TOP 100 *,STOCKNO, STOCKDESC,DNP,NEWNO,SRP,MODELCODE,LOCATION,INVCLASS,subinvclas,onhand FROM PMIS_StockMas WHERE STOCKDESC LIKE '%" & XXX & "%' AND MODELCODE = '" & Left(cboModel, 2) & "'")
        Else
            Set RSTMP = gconDMIS.Execute("SELECT TOP 100 *,STOCKNO, STOCKDESC,DNP,NEWNO,SRP,MODELCODE,LOCATION,INVCLASS,subinvclas,onhand FROM PMIS_StockMas WHERE STOCKDESC LIKE '%" & XXX & "%'")
        End If
    End If
    lsvLIST.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvLIST.ListItems.Add(, , Null2String(RSTMP!STOCKNO))
            ITEM.SubItems(1) = Null2String(RSTMP!STOCKDESC)
            ITEM.SubItems(2) = RSTMP!ID
            
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub
