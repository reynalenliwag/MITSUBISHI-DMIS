VERSION 5.00
Begin VB.Form frmPMISInquiry_PartsInquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PARTS PRICE LOOKUP INQUIRY"
   ClientHeight    =   6645
   ClientLeft      =   315
   ClientTop       =   330
   ClientWidth     =   12735
   ForeColor       =   &H00DEDFDE&
   Icon            =   "PartsInquiry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   12735
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   60
      ScaleHeight     =   375
      ScaleWidth      =   7425
      TabIndex        =   135
      Top             =   90
      Width           =   7425
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
         Left            =   4260
         TabIndex        =   138
         Top             =   60
         Width           =   3105
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
         TabIndex        =   137
         Top             =   60
         Width           =   2655
      End
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
         TabIndex        =   136
         Top             =   60
         Value           =   -1  'True
         Width           =   2445
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   60
      TabIndex        =   17
      Top             =   420
      Width           =   12615
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   0
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   630
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   1
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   960
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   1290
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   1620
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   1950
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   5
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   2280
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   6
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   2610
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   7
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   2940
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   8
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   3270
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   9
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   3600
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   10
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   3930
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   11
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   4260
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   12
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   4590
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   13
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   4920
         Width           =   2295
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
         Left            =   120
         MaxLength       =   30
         TabIndex        =   14
         Text            =   "Text1"
         ToolTipText     =   "Input Valid Part Number"
         Top             =   5250
         Width           =   2295
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
         Left            =   10680
         TabIndex        =   130
         Top             =   630
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   129
         Top             =   990
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   128
         Top             =   1320
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   127
         Top             =   1650
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   126
         Top             =   1980
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   125
         Top             =   2310
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   124
         Top             =   2640
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   123
         Top             =   2970
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   122
         Top             =   3300
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   121
         Top             =   3630
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   120
         Top             =   3960
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   119
         Top             =   4290
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   118
         Top             =   4620
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   117
         Top             =   4950
         Width           =   1845
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
         Left            =   10680
         TabIndex        =   116
         Top             =   5280
         Width           =   1845
      End
      Begin VB.Label Label8 
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
         Height          =   225
         Left            =   10680
         TabIndex        =   115
         Top             =   240
         Width           =   1875
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
         Left            =   9960
         TabIndex        =   114
         Top             =   630
         Width           =   765
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
         Left            =   9960
         TabIndex        =   113
         Top             =   990
         Width           =   765
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
         Left            =   9960
         TabIndex        =   112
         Top             =   1320
         Width           =   765
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
         Left            =   9960
         TabIndex        =   111
         Top             =   1650
         Width           =   765
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
         Left            =   9960
         TabIndex        =   110
         Top             =   1980
         Width           =   765
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
         Left            =   9960
         TabIndex        =   109
         Top             =   2310
         Width           =   765
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
         Left            =   9960
         TabIndex        =   108
         Top             =   2640
         Width           =   765
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
         Left            =   9960
         TabIndex        =   107
         Top             =   2970
         Width           =   765
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
         Left            =   9960
         TabIndex        =   106
         Top             =   3300
         Width           =   765
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
         Left            =   9960
         TabIndex        =   105
         Top             =   3630
         Width           =   765
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
         Left            =   9960
         TabIndex        =   104
         Top             =   3960
         Width           =   765
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
         Left            =   9960
         TabIndex        =   103
         Top             =   4290
         Width           =   765
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
         Left            =   9960
         TabIndex        =   102
         Top             =   4620
         Width           =   765
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
         Left            =   9960
         TabIndex        =   101
         Top             =   4950
         Width           =   765
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
         Left            =   9960
         TabIndex        =   100
         Top             =   5280
         Width           =   765
      End
      Begin VB.Label Label7 
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
         Height          =   375
         Left            =   9810
         TabIndex        =   99
         Top             =   240
         Width           =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   120
         X2              =   12480
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PART NUMBER"
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
         Left            =   120
         TabIndex        =   98
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label5 
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
         Height          =   345
         Left            =   9390
         TabIndex        =   97
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
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
         Left            =   8400
         TabIndex        =   96
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         Height          =   345
         Left            =   6990
         TabIndex        =   95
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   5100
         TabIndex        =   94
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label1 
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
         Height          =   345
         Left            =   2520
         TabIndex        =   93
         Top             =   240
         Width           =   2685
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
         Left            =   9390
         TabIndex        =   92
         Top             =   5280
         Width           =   495
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
         Left            =   8400
         TabIndex        =   91
         Top             =   5280
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   90
         Top             =   5280
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   89
         Top             =   5280
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   88
         Top             =   5280
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   87
         Top             =   4950
         Width           =   495
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
         Left            =   8400
         TabIndex        =   86
         Top             =   4950
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   85
         Top             =   4950
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   84
         Top             =   4950
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   83
         Top             =   4950
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   82
         Top             =   4620
         Width           =   495
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
         Left            =   8400
         TabIndex        =   81
         Top             =   4620
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   80
         Top             =   4620
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   79
         Top             =   4620
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   78
         Top             =   4620
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   77
         Top             =   4290
         Width           =   495
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
         Left            =   8400
         TabIndex        =   76
         Top             =   4290
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   75
         Top             =   4290
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   74
         Top             =   4290
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   73
         Top             =   4290
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   72
         Top             =   3960
         Width           =   495
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
         Left            =   8400
         TabIndex        =   71
         Top             =   3960
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   70
         Top             =   3960
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   69
         Top             =   3960
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   68
         Top             =   3960
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   67
         Top             =   3630
         Width           =   495
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
         Left            =   8400
         TabIndex        =   66
         Top             =   3630
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   65
         Top             =   3630
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   64
         Top             =   3630
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   63
         Top             =   3630
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   62
         Top             =   3300
         Width           =   495
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
         Left            =   8400
         TabIndex        =   61
         Top             =   3300
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   60
         Top             =   3300
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   59
         Top             =   3300
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   58
         Top             =   3300
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   57
         Top             =   2970
         Width           =   495
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
         Left            =   8400
         TabIndex        =   56
         Top             =   2970
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   55
         Top             =   2970
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   54
         Top             =   2970
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   53
         Top             =   2970
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   52
         Top             =   2640
         Width           =   495
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
         Left            =   8400
         TabIndex        =   51
         Top             =   2640
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   50
         Top             =   2640
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   49
         Top             =   2640
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   48
         Top             =   2640
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   47
         Top             =   2310
         Width           =   495
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
         Left            =   8400
         TabIndex        =   46
         Top             =   2310
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   45
         Top             =   2310
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   44
         Top             =   2310
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   43
         Top             =   2310
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   42
         Top             =   1980
         Width           =   495
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
         Left            =   8400
         TabIndex        =   41
         Top             =   1980
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   40
         Top             =   1980
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   39
         Top             =   1980
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   38
         Top             =   1980
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   37
         Top             =   1650
         Width           =   495
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
         Left            =   8400
         TabIndex        =   36
         Top             =   1650
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   35
         Top             =   1650
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   34
         Top             =   1650
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   33
         Top             =   1650
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   32
         Top             =   1320
         Width           =   495
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
         Left            =   8400
         TabIndex        =   31
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   30
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   29
         Top             =   1320
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   28
         Top             =   1320
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   27
         Top             =   990
         Width           =   495
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
         Left            =   8400
         TabIndex        =   26
         Top             =   990
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   25
         Top             =   990
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   24
         Top             =   990
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   23
         Top             =   990
         Width           =   2805
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
         Left            =   9390
         TabIndex        =   22
         Top             =   630
         Width           =   495
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
         Left            =   8400
         TabIndex        =   21
         Top             =   630
         Width           =   945
      End
      Begin VB.Label txtSRP 
         Alignment       =   2  'Center
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
         Left            =   6990
         TabIndex        =   20
         Top             =   630
         Width           =   1395
      End
      Begin VB.Label txtDNPP 
         Alignment       =   1  'Right Justify
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
         Left            =   5100
         TabIndex        =   19
         Top             =   630
         Width           =   2115
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
         Left            =   2520
         TabIndex        =   18
         Top             =   630
         Width           =   2805
      End
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
      Left            =   6005
      TabIndex        =   134
      Top             =   6345
      Width           =   1335
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
      Left            =   5940
      TabIndex        =   133
      Top             =   6150
      Width           =   1335
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
      Left            =   780
      TabIndex        =   132
      Top             =   6150
      Width           =   11835
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
      Left            =   90
      TabIndex        =   131
      Top             =   6150
      Width           =   675
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   270
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   600
      TabIndex        =   15
      Top             =   570
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmPMISInquiry_PartsInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub InquireIT()
    Dim rsDNPP                                         As ADODB.Recordset
    Dim RSPARTMAS                                      As ADODB.Recordset
    Dim SpeakSTOCKNUMBER                               As String
    Dim i                                              As Long
    Dim KIM                                            As Long
    For i = 0 To 14
        If txtPartNo(i).Text <> "" Then
            txtDescrip(i).BorderStyle = 1
            txtDNPP(i).BorderStyle = 1
            txtSRP(i).BorderStyle = 1
            txtModel(i).BorderStyle = 1
            txtICC(i).BorderStyle = 1
            txtSTOCK(i).BorderStyle = 1
            txtLocation(i).BorderStyle = 1

            If Option1.Value = True Then
                Set RSPARTMAS = New ADODB.Recordset
                RSPARTMAS.Open "select STOCKNO,STOCKDESC,dnp,newno,srp,modelcode,location,invclass,subinvclas,onhand from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(txtPartNo(i).Text) & " ", gconDMIS
                If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                    txtDescrip(i).ForeColor = vbBlue: txtDNPP(i).ForeColor = vbBlue: txtSRP(i).ForeColor = vbBlue: txtModel(i).ForeColor = vbBlue
                    txtICC(i).ForeColor = vbBlue: txtSTOCK(i).ForeColor = vbBlue: txtLocation(i).ForeColor = vbBlue
                    txtDescrip(i).Caption = Null2String(RSPARTMAS!STOCKDESC): txtDNPP(i).Caption = Null2String(RSPARTMAS!NEWNO)
                    txtSRP(i).Caption = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP)): txtModel(i).Caption = Null2String(RSPARTMAS!MODELCODE)
                    txtICC(i).Caption = Null2String(RSPARTMAS!InvClass) & Null2String(RSPARTMAS!SubInvClas)
                    txtSTOCK(i).Caption = N2Str2Zero(RSPARTMAS!ONHAND):
                    If N2Str2Zero(RSPARTMAS!ONHAND) > 0 Then
                        txtSTOCK(i).Caption = "Yes"
                    Else
                        txtSTOCK(i).Caption = "N/A"
                    End If
                    txtLocation(i).Caption = Null2String(RSPARTMAS!Location)
                    SpeakSTOCKNUMBER = "": For KIM = 1 To Len(txtPartNo(i).Text): SpeakSTOCKNUMBER = SpeakSTOCKNUMBER & Mid(txtPartNo(i).Text, KIM, 1) & " ": Next

                Else
                    Set rsDNPP = New ADODB.Recordset
                    rsDNPP.Open "select * from PMIS_DNPP where NEWPARTNO=" & N2Str2Null(txtPartNo(i).Text), gconDMIS
                    If Not rsDNPP.EOF And Not rsDNPP.BOF Then
                        txtDescrip(i).ForeColor = vbRed: txtDNPP(i).ForeColor = vbRed: txtSRP(i).ForeColor = vbRed
                        txtModel(i).ForeColor = vbRed: txtICC(i).ForeColor = vbRed: txtSTOCK(i).ForeColor = vbRed: txtLocation(i).ForeColor = vbRed
                        txtDescrip(i).Caption = Null2String(rsDNPP!DESCRIPTIO): txtDNPP(i).Caption = Null2String(rsDNPP!STOCKNUMBER)
                        txtSRP(i).Caption = ToDoubleNumber(N2Str2Zero(rsDNPP!SRP)): txtModel(i).Caption = Null2String(rsDNPP!Model)
                        txtICC(i).Caption = Null2String(rsDNPP!icc): txtSTOCK(i).Caption = "": txtLocation(i).Caption = ""

                    Else
                        Set rsDNPP = New ADODB.Recordset
                        rsDNPP.Open "select * from PMIS_DNPP where PARTNUMBER=" & N2Str2Null(txtPartNo(i).Text), gconDMIS
                        If Not rsDNPP.EOF And Not rsDNPP.BOF Then
                            txtDescrip(i).ForeColor = vbRed: txtDNPP(i).ForeColor = vbRed: txtSRP(i).ForeColor = vbRed
                            txtModel(i).ForeColor = vbRed: txtICC(i).ForeColor = vbRed: txtSTOCK(i).ForeColor = vbRed: txtLocation(i).ForeColor = vbRed
                            txtDescrip(i).Caption = Null2String(rsDNPP!DESCRIPTIO): txtDNPP(i).Caption = Null2String(rsDNPP!NewPARTNO)
                            txtSRP(i).Caption = ToDoubleNumber(N2Str2Zero(rsDNPP!SRP)): txtModel(i).Caption = Null2String(rsDNPP!Model)
                            txtICC(i).Caption = Null2String(rsDNPP!icc): txtSTOCK(i).Caption = "": txtLocation(i).Caption = ""
                        Else
                            txtDescrip(i).Caption = "Not in Master"
                            txtDNPP(i).Caption = "": txtSRP(i).Caption = "": txtModel(i).Caption = ""
                            txtICC(i).Caption = "": txtSTOCK(i).Caption = "": txtLocation(i).Caption = ""
                        End If
                    End If
                End If
            End If

            If Option2.Value = True Then
                Set RSPARTMAS = New ADODB.Recordset
                RSPARTMAS.Open "select STOCKNO,STOCKDESC,dnp,newno,srp,modelcode,location,invclass,subinvclas,onhand from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(txtPartNo(i).Text) & "", gconDMIS
                If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                    txtDescrip(i).ForeColor = vbBlue: txtDNPP(i).ForeColor = vbBlue: txtSRP(i).ForeColor = vbBlue: txtModel(i).ForeColor = vbBlue
                    txtICC(i).ForeColor = vbBlue: txtSTOCK(i).ForeColor = vbBlue: txtLocation(i).ForeColor = vbBlue
                    txtDescrip(i).Caption = Null2String(RSPARTMAS!STOCKDESC): txtDNPP(i).Caption = Null2String(RSPARTMAS!NEWNO)
                    txtSRP(i).Caption = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP)): txtModel(i).Caption = Null2String(RSPARTMAS!MODELCODE)
                    txtICC(i).Caption = Null2String(RSPARTMAS!InvClass) & Null2String(RSPARTMAS!SubInvClas)
                    txtSTOCK(i).Caption = N2Str2Zero(RSPARTMAS!ONHAND): txtLocation(i).Caption = Null2String(RSPARTMAS!Location)
                    SpeakSTOCKNUMBER = "": For KIM = 1 To Len(txtPartNo(i).Text): SpeakSTOCKNUMBER = SpeakSTOCKNUMBER & Mid(txtPartNo(i).Text, KIM, 1) & " ": Next
                Else
                    txtDescrip(i).Caption = "Not in Master"
                    txtDNPP(i).Caption = "": txtSRP(i).Caption = "": txtModel(i).Caption = ""
                    txtICC(i).Caption = "": txtSTOCK(i).Caption = "": txtLocation(i).Caption = ""
                End If
            End If
            If Option3.Value = True Then
                Set rsDNPP = New ADODB.Recordset
                rsDNPP.Open "select * from PMIS_DNPP where PARTNUMBER=" & N2Str2Null(txtPartNo(i).Text), gconDMIS
                If Not rsDNPP.EOF And Not rsDNPP.BOF Then
                    txtDescrip(i).ForeColor = vbRed: txtDNPP(i).ForeColor = vbRed: txtSRP(i).ForeColor = vbRed
                    txtModel(i).ForeColor = vbRed: txtICC(i).ForeColor = vbRed: txtSTOCK(i).ForeColor = vbRed: txtLocation(i).ForeColor = vbRed
                    txtDescrip(i).Caption = Null2String(rsDNPP!DESCRIPTIO): txtDNPP(i).Caption = Null2String(rsDNPP!NewPARTNO)
                    txtSRP(i).Caption = ToDoubleNumber(N2Str2Zero(rsDNPP!SRP)): txtModel(i).Caption = Null2String(rsDNPP!Model)
                    txtICC(i).Caption = Null2String(rsDNPP!icc): txtSTOCK(i).Caption = "": txtLocation(i).Caption = ""
                Else
                    txtDescrip(i).Caption = "Not in Master"
                    txtDNPP(i).Caption = "": txtSRP(i).Caption = "": txtModel(i).Caption = ""
                    txtICC(i).Caption = "": txtSTOCK(i).Caption = "": txtLocation(i).Caption = ""
                End If
            End If
            Call NEW_LogAudit("I", "PARTS AVAILABILITY", "", "", "", "PART NO: " & txtPartNo(i), "", "")
        Else
            txtDescrip(i).Caption = ""
            txtDNPP(i).Caption = ""
            txtSRP(i).Caption = ""
            txtModel(i).Caption = ""
            txtICC(i).Caption = ""
            txtSTOCK(i).Caption = ""
            txtLocation(i).Caption = ""
            txtDescrip(i).BorderStyle = 0
            txtDNPP(i).BorderStyle = 0
            txtSRP(i).BorderStyle = 0
            txtModel(i).BorderStyle = 0
            txtICC(i).BorderStyle = 0
            txtSTOCK(i).BorderStyle = 0
            txtLocation(i).BorderStyle = 0
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
        txtLocation(k).Caption = ""
    Next
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PARTS AVAILABILITY)"
            Call frmALL_AuditInquiry.DisplayHistory("", "PARTS AVAILABILITY", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    initMemvars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISInquiry_PartsInquiry = Nothing
    Unload Me
End Sub

Private Sub txtPartNo_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

