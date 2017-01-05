VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMIOSMonthlyServiceReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthy Service Report"
   ClientHeight    =   11520
   ClientLeft      =   150
   ClientTop       =   -1575
   ClientWidth     =   9660
   FillColor       =   &H8000000D&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "ServiceReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   9660
   Begin VB.Frame Frame2 
      Height          =   8835
      Index           =   3
      Left            =   2370
      TabIndex        =   109
      Top             =   2100
      Width           =   735
      Begin VB.TextBox txtParts 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   63
         Left            =   60
         TabIndex        =   141
         Text            =   "Text1"
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox txtParts 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   62
         Left            =   60
         TabIndex        =   140
         Text            =   "Text1"
         Top             =   150
         Width           =   615
      End
      Begin VB.TextBox txtParts 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   61
         Left            =   60
         TabIndex        =   139
         Text            =   "Text1"
         Top             =   690
         Width           =   615
      End
      Begin VB.TextBox txtParts 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   60
         Left            =   60
         TabIndex        =   138
         Text            =   "Text1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtParts 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   59
         Left            =   60
         TabIndex        =   137
         Text            =   "Text1"
         Top             =   1230
         Width           =   615
      End
      Begin VB.TextBox txtParts 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   58
         Left            =   60
         TabIndex        =   136
         Text            =   "Text1"
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox txtParts 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   57
         Left            =   60
         TabIndex        =   135
         Text            =   "Text1"
         Top             =   1770
         Width           =   615
      End
      Begin VB.TextBox txtParts 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   56
         Left            =   60
         TabIndex        =   134
         Text            =   "Text1"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtParts 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   55
         Left            =   60
         TabIndex        =   133
         Text            =   "Text1"
         Top             =   2310
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   54
         Left            =   60
         TabIndex        =   132
         Text            =   "Text1"
         Top             =   2850
         Width           =   615
      End
      Begin VB.TextBox txtParts 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   60
         TabIndex        =   131
         Text            =   "Text1"
         Top             =   2580
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   52
         Left            =   60
         TabIndex        =   130
         Text            =   "Text1"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   51
         Left            =   60
         TabIndex        =   129
         Text            =   "Text1"
         Top             =   3390
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   60
         TabIndex        =   128
         Text            =   "Text1"
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   49
         Left            =   60
         TabIndex        =   127
         Text            =   "Text1"
         Top             =   3930
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   48
         Left            =   60
         TabIndex        =   126
         Text            =   "Text1"
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   47
         Left            =   60
         TabIndex        =   125
         Text            =   "Text1"
         Top             =   4470
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   46
         Left            =   60
         TabIndex        =   124
         Text            =   "Text1"
         Top             =   4740
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   60
         TabIndex        =   123
         Text            =   "Text1"
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   44
         Left            =   60
         TabIndex        =   122
         Text            =   "Text1"
         Top             =   5010
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   43
         Left            =   60
         TabIndex        =   121
         Text            =   "Text1"
         Top             =   5550
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   42
         Left            =   60
         TabIndex        =   120
         Text            =   "Text1"
         Top             =   5820
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   41
         Left            =   60
         TabIndex        =   119
         Text            =   "Text1"
         Top             =   6090
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   40
         Left            =   60
         TabIndex        =   118
         Text            =   "Text1"
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   60
         TabIndex        =   117
         Text            =   "Text1"
         Top             =   6630
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   38
         Left            =   60
         TabIndex        =   116
         Text            =   "Text1"
         Top             =   7170
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   60
         TabIndex        =   115
         Text            =   "Text1"
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   60
         TabIndex        =   114
         Text            =   "Text1"
         Top             =   7710
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   60
         TabIndex        =   113
         Text            =   "Text1"
         Top             =   7980
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   60
         TabIndex        =   112
         Text            =   "Text1"
         Top             =   8250
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   60
         TabIndex        =   111
         Text            =   "Text1"
         Top             =   6900
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   60
         TabIndex        =   110
         Text            =   "Text1"
         Top             =   8520
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8835
      Index           =   2
      Left            =   1620
      TabIndex        =   76
      Top             =   2100
      Width           =   735
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   60
         TabIndex        =   108
         Text            =   "Text1"
         Top             =   8520
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   60
         TabIndex        =   107
         Text            =   "Text1"
         Top             =   6900
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   60
         TabIndex        =   106
         Text            =   "Text1"
         Top             =   8250
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   60
         TabIndex        =   105
         Text            =   "Text1"
         Top             =   7980
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   60
         TabIndex        =   104
         Text            =   "Text1"
         Top             =   7710
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   60
         TabIndex        =   103
         Text            =   "Text1"
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   60
         TabIndex        =   102
         Text            =   "Text1"
         Top             =   7170
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   60
         TabIndex        =   101
         Text            =   "Text1"
         Top             =   6630
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   60
         TabIndex        =   100
         Text            =   "Text1"
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   60
         TabIndex        =   99
         Text            =   "Text1"
         Top             =   6090
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   60
         TabIndex        =   98
         Text            =   "Text1"
         Top             =   5820
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   60
         TabIndex        =   97
         Text            =   "Text1"
         Top             =   5550
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   60
         TabIndex        =   96
         Text            =   "Text1"
         Top             =   5010
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   60
         TabIndex        =   95
         Text            =   "Text1"
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   60
         TabIndex        =   94
         Text            =   "Text1"
         Top             =   4740
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   60
         TabIndex        =   93
         Text            =   "Text1"
         Top             =   4470
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   60
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   60
         TabIndex        =   91
         Text            =   "Text1"
         Top             =   3930
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   60
         TabIndex        =   90
         Text            =   "Text1"
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   60
         TabIndex        =   89
         Text            =   "Text1"
         Top             =   3390
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   60
         TabIndex        =   88
         Text            =   "Text1"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   60
         TabIndex        =   87
         Text            =   "Text1"
         Top             =   2580
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   60
         TabIndex        =   86
         Text            =   "Text1"
         Top             =   2850
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   60
         TabIndex        =   85
         Text            =   "Text1"
         Top             =   2310
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   84
         Text            =   "Text1"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   60
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   1770
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   81
         Text            =   "Text1"
         Top             =   1230
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   80
         Text            =   "Text1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   690
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   78
         Text            =   "Text1"
         Top             =   150
         Width           =   615
      End
      Begin VB.TextBox txtLabor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8835
      Index           =   1
      Left            =   870
      TabIndex        =   43
      Top             =   2100
      Width           =   735
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   420
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   150
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   73
         Text            =   "Text1"
         Top             =   690
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   71
         Text            =   "Text1"
         Top             =   1230
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   60
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   1770
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   60
         TabIndex        =   67
         Text            =   "Text1"
         Top             =   2310
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   60
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   2850
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   60
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   2580
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   60
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   60
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   3390
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   60
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   60
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   3930
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   60
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   60
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   4470
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   60
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   4740
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   60
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   60
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   5010
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   60
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   5550
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   60
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   5820
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   60
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   6090
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   60
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   60
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   6630
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   60
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   7170
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   60
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   60
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   7710
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   60
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   7980
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   60
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   8250
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   60
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   6900
         Width           =   615
      End
      Begin VB.TextBox txtReleased 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   60
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   8520
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8835
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Top             =   2100
      Width           =   735
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   60
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   8520
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   60
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   6900
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   60
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   8250
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   60
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   7980
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   60
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   7710
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   60
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   60
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   7170
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   60
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   6630
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   60
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   60
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   6090
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   60
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   5820
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   60
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   5550
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   60
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   5010
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   60
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   60
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   4740
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   60
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   4470
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   60
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   60
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   3930
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   60
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   60
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   3390
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   60
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   60
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2580
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   60
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2850
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   60
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2310
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   60
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1770
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1230
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   690
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   150
         Width           =   615
      End
      Begin VB.TextBox txtReceived 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.ComboBox cboYEAR 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1155
   End
   Begin VB.ComboBox cboMONTH 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2265
   End
   Begin Crystal.CrystalReport rptReports 
      Left            =   3540
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CheckBox chkSummary 
      Caption         =   "Summary Only"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1830
      TabIndex        =   2
      Top             =   600
      Width           =   2025
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
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
      Height          =   825
      Left            =   1710
      Picture         =   "ServiceReport.frx":0312
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   990
      Width           =   885
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
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
      Height          =   825
      Left            =   2730
      Picture         =   "ServiceReport.frx":0754
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   990
      Width           =   885
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3330
      TabIndex        =   9
      Top             =   150
      Width           =   735
   End
   Begin VB.Label labRep_or 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   5640
      TabIndex        =   8
      Top             =   1110
      Width           =   4425
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Completed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   10770
      TabIndex        =   7
      Top             =   1140
      Width           =   825
   End
   Begin VB.Label labProgress 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   10200
      TabIndex        =   6
      Top             =   1140
      Width           =   465
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   150
      Width           =   1005
   End
End
Attribute VB_Name = "frmCSMIOSMonthlyServiceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VarDate As String
Dim VarDate_Received As String
Dim VarDate_Released As String
Dim VarLabor As Double
Dim VarParts As Double
Dim VarDiscount As Double
Dim VarOilGas As Double
Dim VarOthers As Double
Dim VarCoJobs As Double
Dim VarSales As Double
Dim VarPDI As Double
Dim VarWarranty As Double
Dim VarSublet As Double
Dim VarTinsmith As Double
Dim VarPainting As Double
Dim VarUndercoat As Double
Dim VarDetailing As Double
Dim VarAircon As Double
Dim VarCounterSales As Double
Dim VarTotal As Double
Dim gconOLDCSMIOS As ADODB.Connection

Private Function OpenOldDb() As Boolean
Dim OLDCSMIOS_Connection As String
With wizVar
     If .VerifyCryptoFile(App.Path & "\CSMIOS.crp") = True Then
        OLDCSMIOS_Connection = .OpenCryptoFile("OLDCSMIOS", "CONNECT")
     End If
End With
On Error Resume Next
deOLDCSMIOS.deConnOLDCSMIOS.Close
On Error GoTo ConnErr
If OLDCSMIOS_Connection <> "" Then
   deOLDCSMIOS.deConnOLDCSMIOS.ConnectionString = OLDCSMIOS_Connection
   Set gconOLDCSMIOS = New Connection
   Set gconOLDCSMIOS = deOLDCSMIOS.deConnOLDCSMIOS
   gconOLDCSMIOS.Open
   OpenOldDb = True
Else
   OpenOldDb = False
End If
Exit Function

ConnErr:
ShowADOErrors gconOLDCSMIOS
End Function

Private Sub cmdPrint_Click()
Dim rsORD_HD As ADODB.Recordset
Dim rsRepor As ADODB.Recordset
Dim rsTdaytran As ADODB.Recordset
Dim rsRO_DET As ADODB.Recordset
Dim rsRR_HD As ADODB.Recordset

VarDate = ""
VarDate_Received = ""
VarDate_Released = ""
VarLabor = 0
VarParts = 0
VarDiscount = 0
VarOilGas = 0
VarOthers = 0
VarCoJobs = 0
VarSales = 0
VarPDI = 0
VarWarranty = 0
VarSublet = 0
VarTinsmith = 0
VarPainting = 0
VarUndercoat = 0
VarDetailing = 0
VarAircon = 0
VarCounterSales = 0
VarTotal = 0
Dim Aldaw As Integer
For Aldaw = 0 To 31
    Set rsRepor = New ADODB.Recordset
    Set rsRepor = gconCSMIOS.Execute("SELECT COUNT(REP_OR) AS Date_Received FROM REPOR WHERE MONTH(DTE_RECD) = " & What_month(cboMONTH.Text) & " AND YEAR(DTE_RECD) = " & cboYEAR.Text & " AND DAY(DTE_RECD) = " & Aldaw)
    If Not rsRepor.EOF And Not rsRepor.BOF Then
       txtReceived(Aldaw).Text = rsRepor![Date_Received]
    End If
Next
Set rsRepor = New ADODB.Recordset
Set rsRepor = gconCSMIOS.Execute("SELECT COUNT(REP_OR) AS Date_Received FROM REPOR WHERE MONTH(DTE_RECD) = " & What_month(cboMONTH.Text) & " AND YEAR(DTE_RECD) = " & cboYEAR.Text)
End Sub

Private Sub Form_Load()
If OpenOldDb = True Then
   CenterMe frmMain, Me, 1
   fillcbomonth cboMONTH
   FillcboYear cboYEAR
End If
End Sub

Sub initMemvars()
VarDate = ""
VarDate_Received = ""
VarDate_Released = ""
VarLabor = 0
VarParts = 0
VarDiscount = 0
VarOilGas = 0
VarOthers = 0
VarCoJobs = 0
VarSales = 0
VarPDI = 0
VarWarranty = 0
VarSublet = 0
VarTinsmith = 0
VarPainting = 0
VarUndercoat = 0
VarDetailing = 0
VarAircon = 0
VarCounterSales = 0
VarTotal = 0
End Sub
