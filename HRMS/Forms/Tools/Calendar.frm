VERSION 5.00
Begin VB.Form frmToolsCalendar 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HRMS Calendar"
   ClientHeight    =   4875
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4650
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Calendar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Calendar.frx":0442
   ScaleHeight     =   4875
   ScaleWidth      =   4650
   Begin VB.CommandButton frmdataclk 
      Caption         =   "Command1"
      Height          =   195
      Left            =   4440
      TabIndex        =   49
      Top             =   90
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   120
      Top             =   4320
   End
   Begin VB.CommandButton cmdLY 
      BackColor       =   &H00DEDFDE&
      Caption         =   "<"
      Height          =   255
      Left            =   1620
      MouseIcon       =   "Calendar.frx":317E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   90
      Width           =   645
   End
   Begin VB.CommandButton cmdRY 
      BackColor       =   &H00DEDFDE&
      Caption         =   ">"
      Height          =   255
      Left            =   2430
      MouseIcon       =   "Calendar.frx":32D0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   90
      Width           =   645
   End
   Begin VB.CommandButton cmdRight 
      BackColor       =   &H00DEDFDE&
      Caption         =   ">"
      Height          =   255
      Left            =   900
      MouseIcon       =   "Calendar.frx":3422
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   90
      Width           =   645
   End
   Begin VB.CommandButton cmdLeft 
      BackColor       =   &H00DEDFDE&
      Caption         =   "<"
      Height          =   255
      Left            =   90
      MouseIcon       =   "Calendar.frx":3574
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   90
      Width           =   645
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   4455
      TabIndex        =   50
      Top             =   870
      Width           =   4455
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   57
         Top             =   90
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   690
         TabIndex        =   56
         Top             =   90
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1380
         TabIndex        =   55
         Top             =   90
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2010
         TabIndex        =   54
         Top             =   90
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Thurs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2640
         TabIndex        =   53
         Top             =   90
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   3330
         TabIndex        =   52
         Top             =   90
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   3960
         TabIndex        =   51
         Top             =   90
         Width           =   375
      End
   End
   Begin VB.Label lbltime 
      BackColor       =   &H00DEDFDE&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3210
      TabIndex        =   48
      Top             =   420
      Width           =   1335
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   35
      Left            =   90
      TabIndex        =   47
      Top             =   4290
      Width           =   495
   End
   Begin VB.Label lblYear 
      BackColor       =   &H00DEDFDE&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1620
      TabIndex        =   42
      Top             =   420
      Width           =   1455
   End
   Begin VB.Label lblMonth 
      BackColor       =   &H00DEDFDE&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   41
      Top             =   420
      Width           =   1425
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   41
      Left            =   4050
      TabIndex        =   40
      Top             =   4290
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   40
      Left            =   3390
      TabIndex        =   39
      Top             =   4290
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   39
      Left            =   2730
      TabIndex        =   38
      Top             =   4290
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   38
      Left            =   2070
      TabIndex        =   37
      Top             =   4290
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   37
      Left            =   1410
      TabIndex        =   36
      Top             =   4290
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   36
      Left            =   750
      TabIndex        =   35
      Top             =   4290
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   34
      Left            =   4050
      TabIndex        =   34
      Top             =   3690
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   33
      Left            =   3390
      TabIndex        =   33
      Top             =   3690
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   32
      Left            =   2730
      TabIndex        =   32
      Top             =   3690
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   31
      Left            =   2070
      TabIndex        =   31
      Top             =   3690
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   30
      Left            =   1410
      TabIndex        =   30
      Top             =   3690
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   29
      Left            =   750
      TabIndex        =   29
      Top             =   3690
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   28
      Left            =   90
      TabIndex        =   28
      Top             =   3690
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   27
      Left            =   4050
      TabIndex        =   27
      Top             =   3090
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   26
      Left            =   3390
      TabIndex        =   26
      Top             =   3090
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   25
      Left            =   2730
      TabIndex        =   25
      Top             =   3090
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   24
      Left            =   2070
      TabIndex        =   24
      Top             =   3090
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   23
      Left            =   1410
      TabIndex        =   23
      Top             =   3090
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   22
      Left            =   750
      TabIndex        =   22
      Top             =   3090
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   21
      Left            =   90
      TabIndex        =   21
      Top             =   3090
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   20
      Left            =   4050
      TabIndex        =   20
      Top             =   2490
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   19
      Left            =   3390
      TabIndex        =   19
      Top             =   2490
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   18
      Left            =   2730
      TabIndex        =   18
      Top             =   2490
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   17
      Left            =   2070
      TabIndex        =   17
      Top             =   2490
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   16
      Left            =   1410
      TabIndex        =   16
      Top             =   2490
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   15
      Left            =   750
      TabIndex        =   15
      Top             =   2490
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   14
      Left            =   90
      TabIndex        =   14
      Top             =   2490
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   13
      Left            =   4050
      TabIndex        =   13
      Top             =   1890
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   12
      Left            =   3390
      TabIndex        =   12
      Top             =   1890
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   11
      Left            =   2730
      TabIndex        =   11
      Top             =   1890
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   10
      Left            =   2070
      TabIndex        =   10
      Top             =   1890
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   9
      Left            =   1410
      TabIndex        =   9
      Top             =   1890
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   8
      Left            =   750
      TabIndex        =   8
      Top             =   1890
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   7
      Left            =   90
      TabIndex        =   7
      Top             =   1890
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   4050
      TabIndex        =   6
      Top             =   1290
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   3390
      TabIndex        =   5
      Top             =   1290
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   2730
      TabIndex        =   4
      Top             =   1290
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   2070
      TabIndex        =   3
      Top             =   1290
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   1410
      TabIndex        =   2
      Top             =   1290
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   750
      TabIndex        =   1
      Top             =   1290
      Width           =   495
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   1290
      Width           =   495
   End
End
Attribute VB_Name = "frmToolsCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curday As Integer
Dim mon As Integer
Dim yr As Integer
Dim dayi As Integer
Dim dayinyear As Integer, a As String, b As Integer, x As Integer 'dims form varaibles
Dim dayx As Long, dayq As Integer ' other variables that are global diminished in module
Dim tstcurday(365) As Integer, tstmon(365) As Integer, tstyr(365) As Integer, c(365) As String
Dim fL As Boolean, numentries As Integer, numforvar As Integer, q As Integer
Private Sub cmdLeft_Click()
dayi = 0
dayq = 0
If yr > 1 Then  'if year is valid then
    If mon > 1 Then 'if month is valid then
        mon = mon - 1   'make month previous month
    Else    'if month is not valid (previous january)
        yr = yr - 1 'make year prvious year
        mon = 12    'set the month to the last month of the year
    End If
Else
    MsgBox ("This Calendar can only go into A.D. years")
End If
code   'goto code
End Sub
Private Sub cmdLY_Click()
dayi = 0
dayq = 0
If yr > 1 Then  'if year is valid then
    yr = yr - 1 'make year = last year
Else
    MsgBox ("This Calendar can only go into A.D. years")
End If
code ' goto code
End Sub
Private Sub cmdRight_Click()
dayi = 0
dayq = 0
If mon < 12 Then    ' if month is not december then
    mon = mon + 1   'month = next month
Else    'if month is december then
    yr = yr + 1 'goto next year
    mon = 1 'month = january
End If
code 'goto code
End Sub
Private Sub cmdRY_Click()
dayq = 0
dayi = 0
yr = yr + 1 ' goto next year (no extra code cuz year can be up to infinity
code    'goto code
End Sub
Private Sub Form_Load()
CenterMe frmMain, Me, 1
dayi = 0
yr = Year(LOGDATE)     'make the year = the year set in the users date/time properties
mon = Month(LOGDATE)   'make the month=the month set in the users date/time properties
code    'goto code with current year/month
End Sub

Public Sub code()
Dim monthx(1 To 12) 'dim the 12 months
dayx = 6    'start off on the right foot(Jan 0001 was the 7th day so 6 before it)
For x = 0 To 41
    lblDate(x) = ""     'clear all date labels
Next x
monthx(1) = 31  'set defaut for jan to 31
monthx(2) = 28  '...
monthx(3) = 31 '...
monthx(4) = 30 '...
monthx(5) = 31 '.
monthx(6) = 30 '...
monthx(7) = 31 '...
monthx(8) = 31 '...
monthx(9) = 30 '...
monthx(10) = 31 '...
monthx(11) = 30 '...
monthx(12) = 31 '...
If mon = 1 Then lblMonth = "January"        'tells label what month it is
If mon = 2 Then lblMonth = "Ferbuary"       '...
If mon = 3 Then lblMonth = "March"
If mon = 4 Then lblMonth = "April"
If mon = 5 Then lblMonth = "May"
If mon = 6 Then lblMonth = "June"
If mon = 7 Then lblMonth = "July"
If mon = 8 Then lblMonth = "August"
If mon = 9 Then lblMonth = "September"
If mon = 10 Then lblMonth = "October"
If mon = 11 Then lblMonth = "November"
If mon = 12 Then lblMonth = "December" '...
lblYear = yr            'displays year calandar is in onto year label
If yr / 4 = Int(yr / 4) Then monthx(2) = 29 'if possible leap year then set feb to 29 dias
If yr >= 1752 Then                          'if gregarian reform has taken place (new leap year rules) then
    If yr / 100 = Int(yr / 100) Then monthx(2) = 28 'if a year is divisible by 100 and not 400 then not leap year
    If yr / 400 = Int(yr / 400) Then monthx(2) = 29 'if true than is leap year
Else    'if gregarian reform has not taken place(old rules previous 1752)
    If yr / 4 = Int(yr / 4) Then monthx(2) = 29 Else monthx(2) = 28 'if leap year .. else
End If
'\/ adds up days until the selected year

For x = 1 To yr - 1 'does not go into current year because current year is not full
    dayinyear = 365 'set norm days in a year
    If x >= 1752 Then
        If x / 4 = Int(x / 4) Then dayinyear = 366 'leapyear
        If x / 100 = Int(x / 100) Then dayinyear = 365  'set days in the year
        If x / 400 = Int(x / 400) Then dayinyear = 366  'set days in the year
    Else
        dayinyear = 365 'set days in year to norm (365 days)
        If x / 4 = Int(x / 4) Then dayinyear = 366 'if leap year then (addday-366)
    End If
    If x = 1752 Then dayinyear = 355 ' only 355 days in yr 1752 because gregarian reform
    dayx = dayx + dayinyear 'dayx keeps total of dias until select yr
Next x
If yr = 1752 Then monthx(9) = 19    'if user selects a month greater than 9 in year 1752 then must add correct amount of days to month 9(september)
For x = 1 To (mon - 1)  'add up days in current year up until selected month
    dayx = dayx + monthx(x)
Next x
Do While dayx > 7   'subtracts week everytime
    dayx = dayx - 7 'sees what day to start on
Loop
a = 7
b = dayx
For x = 1 To 40
    lblDate(x - 1).BorderStyle = 0
Next x
If monthx(mon) = 19 Then 'if month where gregarian reform took place
    For x = 1 To 3
        lblDate(x + 1) = x
        lblDate(x + 1).BorderStyle = 1
    Next x
    For x = 14 To 30
        lblDate(x - 10) = x
        lblDate(x - 10).BorderStyle = 1
    Next x
Else    'if other then
    If b = 7 Then b = 0
    For x = 1 To monthx(mon)
        lblDate(b) = x
        lblDate(b).BorderStyle = 1
        b = b + 1
    Next x
End If
'If lblDate(35).BorderStyle = 0 Then
'    Me.WindowState = 0
'    Me.Height = 6250
'    Me.Width = 7600
'Else
'    Me.WindowState = 0
'    Me.Height = 7065
'    Me.Width = 7600
'End If
For x = 0 To 41
    lblDate(x).BackColor = &H8000000A
Next x
dayq = Day(LOGDATE)
'highlights day if it is today
If mon = Month(LOGDATE) Then
    If yr = Year(LOGDATE) Then
        For x = 0 To 41
            If numericval(lblDate(x).Caption) = dayq Then
                lblDate(x).BackColor = RGB(250, 0, 0)
            End If
        Next x
    End If
End If
'\/ sets where to put time box
'If lblDate(1).Caption = "" Then
    'lbltime.Top = 2040
    'lbltime.Left = 360
'Else
 '   If lblDate(33).Caption = "" Then
        'lbltime.Top = 4900
        'lbltime.Left = 5600
'    Else
        'Me.WindowState = 0
        'Me.Height = 7035
        'lbltime.Top = 5600
        'lbltime.Left = 5600
'    End If
'End If
'######################################################
'######################################################
lblDate(40).Visible = False
lblDate(41).Visible = False
Timer1.Enabled = True
'###################

'#####################
fL = False
End Sub

Private Sub mnuGoToDate_Click()
dmydate = InputBox("Enter the date in the format mm/dd/yyyy", "Input Date")
dmydate = Trim(dmydate)
If Len(dmydate) = 10 Then
    If Left(dmydate, 2) > 0 Then
        If Left(dmydate, 2) < 13 Then
            If Right(dmydate, 4) > 0 Then
                mon = Left(dmydate, 2)
                yr = Right(dmydate, 4)
                dayi = Mid(dmydate, 4, 2)
                
                code
            End If
        End If
    End If
End If
If mon <> Left(dmydate, 2) Then
    If yr <> Right(dmydate, 4) Then MsgBox ("Enter Date in Correct form!")
End If
End Sub

Private Sub redoyear()
Me.SetFocus
End Sub

Private Sub mnuPrintCal_Click()
lbltime.Visible = False
PrintForm
lbltime.Visible = True
End Sub

Private Sub Timer1_Timer()
lbltime = Time
End Sub
