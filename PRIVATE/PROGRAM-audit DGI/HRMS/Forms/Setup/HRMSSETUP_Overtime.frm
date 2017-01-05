VERSION 5.00
Begin VB.Form frmSETUP_Overtime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Overtime Setup"
   ClientHeight    =   6225
   ClientLeft      =   585
   ClientTop       =   810
   ClientWidth     =   9765
   ForeColor       =   &H8000000F&
   Icon            =   "HRMSSETUP_Overtime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   9765
   Begin VB.PictureBox Picture2 
      Height          =   345
      Left            =   4920
      ScaleHeight     =   285
      ScaleWidth      =   4725
      TabIndex        =   96
      Top             =   450
      Width           =   4785
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "Pay Code"
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
         Left            =   1920
         TabIndex        =   98
         Top             =   30
         Width           =   945
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "Pay Rate"
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
         Left            =   3270
         TabIndex        =   97
         Top             =   30
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Left            =   90
      ScaleHeight     =   285
      ScaleWidth      =   4725
      TabIndex        =   93
      Top             =   450
      Width           =   4785
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "Pay Rate"
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
         Left            =   3270
         TabIndex        =   95
         Top             =   30
         Width           =   1335
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "Pay Code"
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
         Left            =   1920
         TabIndex        =   94
         Top             =   30
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   4920
      TabIndex        =   47
      Top             =   720
      Width           =   4785
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   11
         Left            =   1980
         TabIndex        =   62
         Top             =   630
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   11
         Left            =   2850
         TabIndex        =   61
         Top             =   630
         Width           =   345
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   11
         Left            =   3270
         TabIndex        =   60
         Text            =   "0.00"
         Top             =   630
         Width           =   1425
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   12
         Left            =   1980
         TabIndex        =   59
         Top             =   990
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   12
         Left            =   2850
         TabIndex        =   58
         Top             =   990
         Width           =   345
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   12
         Left            =   3270
         TabIndex        =   57
         Text            =   "0.00"
         Top             =   990
         Width           =   1425
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   13
         Left            =   1980
         TabIndex        =   56
         Top             =   1350
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   13
         Left            =   2850
         TabIndex        =   55
         Top             =   1350
         Width           =   345
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   13
         Left            =   3270
         TabIndex        =   54
         Text            =   "0.00"
         Top             =   1350
         Width           =   1425
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   14
         Left            =   1980
         TabIndex        =   53
         Top             =   1710
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   14
         Left            =   2850
         TabIndex        =   52
         Top             =   1710
         Width           =   345
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   14
         Left            =   3270
         TabIndex        =   51
         Text            =   "0.00"
         Top             =   1710
         Width           =   1425
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   15
         Left            =   1980
         TabIndex        =   50
         Top             =   2070
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   15
         Left            =   2850
         TabIndex        =   49
         Top             =   2070
         Width           =   345
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   15
         Left            =   3270
         TabIndex        =   48
         Text            =   "0.00"
         Top             =   2070
         Width           =   1425
      End
      Begin VB.Label Label19 
         Caption         =   "Regular OT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   68
         Top             =   330
         Width           =   1995
      End
      Begin VB.Label Label18 
         Caption         =   "Reg. Holiday"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   67
         Top             =   690
         Width           =   1995
      End
      Begin VB.Label Label17 
         Caption         =   "Special Hol."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   66
         Top             =   1050
         Width           =   1995
      End
      Begin VB.Label Label16 
         Caption         =   "Day Off"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   65
         Top             =   1410
         Width           =   1995
      End
      Begin VB.Label Label15 
         Caption         =   "Day Off/Reg. Hol."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   64
         Top             =   1770
         Width           =   1995
      End
      Begin VB.Label Label14 
         Caption         =   "Day Off/ Spe. Hol."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   63
         Top             =   2100
         Width           =   1995
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2535
      Left            =   90
      TabIndex        =   22
      Top             =   3600
      Width           =   4785
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   8
         Left            =   1980
         TabIndex        =   106
         Top             =   1350
         Width           =   825
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   5
         Left            =   1980
         TabIndex        =   39
         Top             =   270
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   5
         Left            =   2850
         TabIndex        =   38
         Top             =   270
         Width           =   345
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   5
         Left            =   3270
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   270
         Width           =   1425
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   6
         Left            =   1980
         TabIndex        =   36
         Top             =   630
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   6
         Left            =   2850
         TabIndex        =   35
         Top             =   630
         Width           =   345
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   6
         Left            =   3270
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   630
         Width           =   1425
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   7
         Left            =   1980
         TabIndex        =   33
         Top             =   990
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   7
         Left            =   2850
         TabIndex        =   32
         Top             =   990
         Width           =   345
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   7
         Left            =   3270
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   990
         Width           =   1425
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   8
         Left            =   2850
         TabIndex        =   30
         Top             =   1350
         Width           =   345
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   8
         Left            =   3270
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   1350
         Width           =   1425
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   9
         Left            =   1980
         TabIndex        =   28
         Top             =   1710
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   9
         Left            =   2850
         TabIndex        =   27
         Top             =   1710
         Width           =   345
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   9
         Left            =   3270
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   1710
         Width           =   1425
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   10
         Left            =   1980
         TabIndex        =   25
         Top             =   2070
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   10
         Left            =   2850
         TabIndex        =   24
         Top             =   2070
         Width           =   345
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   10
         Left            =   3270
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   2070
         Width           =   1425
      End
      Begin VB.Label Label12 
         Caption         =   "Regular ND"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   45
         Top             =   330
         Width           =   1995
      End
      Begin VB.Label Label11 
         Caption         =   "Reg. Holiday"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   44
         Top             =   690
         Width           =   1995
      End
      Begin VB.Label Label10 
         Caption         =   "Special Hol."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   43
         Top             =   1050
         Width           =   1995
      End
      Begin VB.Label Label9 
         Caption         =   "Day Off"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   42
         Top             =   1410
         Width           =   1995
      End
      Begin VB.Label Label8 
         Caption         =   "Day Off/Reg. Hol."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   41
         Top             =   1770
         Width           =   1995
      End
      Begin VB.Label Label7 
         Caption         =   "Day Off/ Spe. Hol."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   40
         Top             =   2100
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   90
      TabIndex        =   0
      Top             =   720
      Width           =   4785
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   3270
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   2070
         Width           =   1425
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   4
         Left            =   2850
         TabIndex        =   20
         Top             =   2070
         Width           =   345
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   1980
         TabIndex        =   19
         Top             =   2070
         Width           =   825
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   3270
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   1710
         Width           =   1425
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   3
         Left            =   2850
         TabIndex        =   17
         Top             =   1710
         Width           =   345
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   1980
         TabIndex        =   16
         Top             =   1710
         Width           =   825
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   3270
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   1350
         Width           =   1425
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   2
         Left            =   2850
         TabIndex        =   14
         Top             =   1350
         Width           =   345
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   1980
         TabIndex        =   13
         Top             =   1350
         Width           =   825
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   3270
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   990
         Width           =   1425
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   2850
         TabIndex        =   11
         Top             =   990
         Width           =   345
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   1980
         TabIndex        =   10
         Top             =   990
         Width           =   825
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   3270
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   630
         Width           =   1425
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   2850
         TabIndex        =   8
         Top             =   630
         Width           =   345
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   1980
         TabIndex        =   7
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label6 
         Caption         =   "Day Off/ Spe. Hol."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   2100
         Width           =   1995
      End
      Begin VB.Label Label5 
         Caption         =   "Day Off/Reg. Hol."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   1770
         Width           =   1995
      End
      Begin VB.Label Label4 
         Caption         =   "Day Off"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   1410
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Special Hol."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   3
         Top             =   1050
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "Reg. Holiday"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   2
         Top             =   690
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Regular OT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   1995
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2925
      Left            =   4920
      TabIndex        =   69
      Top             =   3210
      Width           =   4785
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   22
         Left            =   1980
         TabIndex        =   104
         Top             =   2430
         Width           =   825
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   22
         Left            =   2850
         TabIndex        =   103
         Top             =   2430
         Width           =   345
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   22
         Left            =   3270
         TabIndex        =   102
         Text            =   "0.00"
         Top             =   2430
         Width           =   1425
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   21
         Left            =   3270
         TabIndex        =   87
         Text            =   "0.00"
         Top             =   2070
         Width           =   1425
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   21
         Left            =   2850
         TabIndex        =   86
         Top             =   2070
         Width           =   345
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   21
         Left            =   1980
         TabIndex        =   85
         Top             =   2070
         Width           =   825
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   20
         Left            =   3270
         TabIndex        =   84
         Text            =   "0.00"
         Top             =   1710
         Width           =   1425
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   20
         Left            =   2850
         TabIndex        =   83
         Top             =   1710
         Width           =   345
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   20
         Left            =   1980
         TabIndex        =   82
         Top             =   1710
         Width           =   825
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   19
         Left            =   3270
         TabIndex        =   81
         Text            =   "0.00"
         Top             =   1350
         Width           =   1425
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   19
         Left            =   2850
         TabIndex        =   80
         Top             =   1350
         Width           =   345
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   19
         Left            =   1980
         TabIndex        =   79
         Top             =   1350
         Width           =   825
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   18
         Left            =   3270
         TabIndex        =   78
         Text            =   "0.00"
         Top             =   990
         Width           =   1425
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   18
         Left            =   2850
         TabIndex        =   77
         Top             =   990
         Width           =   345
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   18
         Left            =   1980
         TabIndex        =   76
         Top             =   990
         Width           =   825
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   17
         Left            =   3270
         TabIndex        =   75
         Text            =   "0.00"
         Top             =   630
         Width           =   1425
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   17
         Left            =   2850
         TabIndex        =   74
         Top             =   630
         Width           =   345
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   17
         Left            =   1980
         TabIndex        =   73
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Index           =   16
         Left            =   3270
         TabIndex        =   72
         Text            =   "0.00"
         Top             =   270
         Width           =   1425
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   315
         Index           =   16
         Left            =   2850
         TabIndex        =   71
         Top             =   270
         Width           =   345
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
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
         Index           =   16
         Left            =   1980
         TabIndex        =   70
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label32 
         Caption         =   "Other OT7"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   105
         Top             =   2460
         Width           =   1995
      End
      Begin VB.Label Label25 
         Caption         =   "Other OT6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   100
         Top             =   2100
         Width           =   1995
      End
      Begin VB.Label Label24 
         Caption         =   "Other OT5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   99
         Top             =   1770
         Width           =   1995
      End
      Begin VB.Label Label23 
         Caption         =   "Other OT4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   91
         Top             =   1410
         Width           =   1995
      End
      Begin VB.Label Label22 
         Caption         =   "Other OT3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   90
         Top             =   1050
         Width           =   1995
      End
      Begin VB.Label Label21 
         Caption         =   "Other OT2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   89
         Top             =   690
         Width           =   1995
      End
      Begin VB.Label Label20 
         Caption         =   "Other OT1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   88
         Top             =   330
         Width           =   1995
      End
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Night Differential"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   101
      Top             =   3300
      Width           =   4785
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Excess of 8 Hours"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   92
      Top             =   120
      Width           =   4785
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "First 8 Hours"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   46
      Top             =   120
      Width           =   4785
   End
End
Attribute VB_Name = "frmSETUP_Overtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOTSetup                                As ADODB.Recordset

Private Sub cmd_Click(INDEX As Integer)
    frmHRMSOvertimeCodes.lblINDEX.Caption = INDEX
    frmHRMSOvertimeCodes.Show 1
    
    txtCode(INDEX).Text = OVERTIME_CODES
    txtRate(INDEX).Text = OVERTIME_RATE
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
    RsRefresh
    STOREMEMVARS
End Sub

Sub RsRefresh()
    Set rsOTSetup = New ADODB.Recordset
    Set rsOTSetup = gconDMIS.Execute("Select * from HRMS_OTSetup")
End Sub

Sub STOREMEMVARS()
    If Not rsOTSetup.EOF And Not rsOTSetup.BOF Then
        txtCode(0).Text = Null2String(rsOTSetup!Code)
        txtCode(1).Text = Null2String(rsOTSetup!Code1)
        txtCode(2).Text = Null2String(rsOTSetup!Code2)
        txtCode(3).Text = Null2String(rsOTSetup!Code3)
        txtCode(4).Text = Null2String(rsOTSetup!Code4)
        txtCode(5).Text = Null2String(rsOTSetup!Code5)
        txtCode(6).Text = Null2String(rsOTSetup!Code6)
        txtCode(7).Text = Null2String(rsOTSetup!Code7)
        txtCode(8).Text = Null2String(rsOTSetup!Code8)
        txtCode(9).Text = Null2String(rsOTSetup!Code9)
        txtCode(10).Text = Null2String(rsOTSetup!Code10)
        txtCode(11).Text = Null2String(rsOTSetup!Code11)
        txtCode(12).Text = Null2String(rsOTSetup!Code12)
        txtCode(13).Text = Null2String(rsOTSetup!Code13)
        txtCode(14).Text = Null2String(rsOTSetup!Code14)
        txtCode(15).Text = Null2String(rsOTSetup!Code15)
        txtCode(16).Text = Null2String(rsOTSetup!Code16)
        txtCode(17).Text = Null2String(rsOTSetup!Code17)
        txtCode(18).Text = Null2String(rsOTSetup!Code18)
        txtCode(19).Text = Null2String(rsOTSetup!Code19)
        txtCode(20).Text = Null2String(rsOTSetup!Code20)
        txtCode(21).Text = Null2String(rsOTSetup!Code21)
        txtCode(22).Text = Null2String(rsOTSetup!Code22)

        txtRate(0).Text = N2Str2Zero(rsOTSetup!Rate)
        txtRate(1).Text = N2Str2Zero(rsOTSetup!Rate1)
        txtRate(2).Text = N2Str2Zero(rsOTSetup!Rate2)
        txtRate(3).Text = N2Str2Zero(rsOTSetup!Rate3)
        txtRate(4).Text = N2Str2Zero(rsOTSetup!Rate4)
        txtRate(5).Text = N2Str2Zero(rsOTSetup!Rate5)
        txtRate(6).Text = N2Str2Zero(rsOTSetup!Rate6)
        txtRate(7).Text = N2Str2Zero(rsOTSetup!Rate7)
        txtRate(8).Text = N2Str2Zero(rsOTSetup!Rate8)
        txtRate(9).Text = N2Str2Zero(rsOTSetup!Rate9)
        txtRate(10).Text = N2Str2Zero(rsOTSetup!Rate10)
        txtRate(11).Text = N2Str2Zero(rsOTSetup!Rate11)
        txtRate(12).Text = N2Str2Zero(rsOTSetup!Rate12)
        txtRate(13).Text = N2Str2Zero(rsOTSetup!Rate13)
        txtRate(14).Text = N2Str2Zero(rsOTSetup!Rate14)
        txtRate(15).Text = N2Str2Zero(rsOTSetup!Rate15)
        txtRate(16).Text = N2Str2Zero(rsOTSetup!Rate16)
        txtRate(17).Text = N2Str2Zero(rsOTSetup!Rate17)
        txtRate(18).Text = N2Str2Zero(rsOTSetup!Rate18)
        txtRate(19).Text = N2Str2Zero(rsOTSetup!Rate19)
        txtRate(20).Text = N2Str2Zero(rsOTSetup!Rate20)
        txtRate(21).Text = N2Str2Zero(rsOTSetup!Rate21)
        txtRate(22).Text = N2Str2Zero(rsOTSetup!Rate22)
    End If
End Sub

