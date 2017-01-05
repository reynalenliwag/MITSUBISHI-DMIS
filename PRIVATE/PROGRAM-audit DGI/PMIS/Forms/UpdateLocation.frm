VERSION 5.00
Begin VB.Form frmPMISUpdateLocation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Parts Location"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   ForeColor       =   &H00E0E0E0&
   Icon            =   "UpdateLocation.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   8145
   Begin VB.CommandButton cmdCancel 
      Caption         =   "cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   7320
      MouseIcon       =   "UpdateLocation.frx":0152
      MousePointer    =   99  'Custom
      Picture         =   "UpdateLocation.frx":02A4
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Cancel"
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6600
      MouseIcon       =   "UpdateLocation.frx":05E2
      MousePointer    =   99  'Custom
      Picture         =   "UpdateLocation.frx":0734
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Save Changes"
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox txtNewLocation15 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      MouseIcon       =   "UpdateLocation.frx":0A84
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   5820
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation15 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   61
      Top             =   5820
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo15 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   14
      Top             =   5820
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation14 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   28
      Top             =   5430
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation14 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   59
      Top             =   5430
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo14 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   13
      Top             =   5430
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation13 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   27
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation13 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   57
      Top             =   5040
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo13 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   12
      Top             =   5040
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   26
      Top             =   4650
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation12 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   55
      Top             =   4650
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   11
      Top             =   4650
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation11 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   25
      Top             =   4260
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation11 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   53
      Top             =   4260
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo11 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   10
      Top             =   4260
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   24
      Top             =   3870
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation10 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   51
      Top             =   3870
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   9
      Top             =   3870
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   23
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation9 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   49
      Top             =   3480
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   8
      Top             =   3480
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   22
      Top             =   3090
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation8 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   47
      Top             =   3090
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   7
      Top             =   3090
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   21
      Top             =   2700
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation7 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   45
      Top             =   2700
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   6
      Top             =   2700
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   20
      Top             =   2310
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation6 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   43
      Top             =   2310
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   5
      Top             =   2310
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   19
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation5 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   41
      Top             =   1920
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   4
      Top             =   1920
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   18
      Top             =   1530
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation4 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   39
      Top             =   1530
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1530
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   17
      Top             =   1140
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation3 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   37
      Top             =   1140
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1140
      Width           =   2115
   End
   Begin VB.TextBox txtNewLocation2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   16
      Top             =   750
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation2 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   35
      Top             =   750
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   1
      Top             =   750
      Width           =   2145
   End
   Begin VB.TextBox txtNewLocation1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5160
      MaxLength       =   100
      TabIndex        =   15
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtOldLocation1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3150
      MaxLength       =   100
      TabIndex        =   31
      Top             =   360
      Width           =   1995
   End
   Begin VB.TextBox txtPartNo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   990
      MaxLength       =   100
      TabIndex        =   0
      Top             =   360
      Width           =   2115
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   60
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   58
      Top             =   5490
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   56
      Top             =   5100
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   54
      Top             =   4710
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   52
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   50
      Top             =   3930
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   48
      Top             =   3540
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   46
      Top             =   3150
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   44
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   42
      Top             =   2370
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   40
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   38
      Top             =   1590
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   36
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   34
      Top             =   810
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "New Location"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5130
      TabIndex        =   33
      Top             =   60
      Width           =   1965
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Location"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3150
      TabIndex        =   32
      Top             =   60
      Width           =   1965
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   30
      Top             =   420
      Width           =   1335
   End
End
Attribute VB_Name = "frmPMISUpdateLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function CheckLocationAssigned(CheckSTOCKNO As String, CheckLocation As String) As String
    Dim rsCHECKPARTMAS                                 As ADODB.Recordset
    Set rsCHECKPARTMAS = New ADODB.Recordset
    Set rsCHECKPARTMAS = gconDMIS.Execute("Select STOCKNO from PMIS_PARTMAS where LOCATION = '" & CheckLocation & "' and STOCKNO <> '" & CheckSTOCKNO & "'")
    If Not rsCHECKPARTMAS.EOF And Not rsCHECKPARTMAS.BOF Then
        CheckLocationAssigned = Null2String(rsCHECKPARTMAS!STOCKNO)
    Else
        CheckLocationAssigned = "NA"
    End If
End Function

Sub SetLocation()
    Dim RSPARTMAS                                      As ADODB.Recordset
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo1.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation1.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo2.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation2.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo3.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation3.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo4.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation4.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo5.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation5.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo6.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation6.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo7.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation7.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo8.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation8.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo9.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation9.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo10.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation10.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo11.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation11.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo12.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation12.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo13.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation13.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo14.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation14.Text = Null2String(RSPARTMAS!Location)
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select LOCATION from PMIS_PARTMAS Where STOCKNO = " & N2Str2Null(txtPartNo15.Text))
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then txtOldLocation15.Text = Null2String(RSPARTMAS!Location)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo Errorcode:

    If MsgBox("Update Parts Master File with these New Location?", vbQuestion + vbYesNo, "Confirm Update...") = vbYes Then
        Screen.MousePointer = 11
        If Trim(txtPartNo1.Text) <> "" And Trim(txtNewLocation1.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation1.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo1.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo2.Text) <> "" And Trim(txtNewLocation2.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation2.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo2.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo3.Text) <> "" And Trim(txtNewLocation3.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation3.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo3.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo4.Text) <> "" And Trim(txtNewLocation4.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation4.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo4.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo5.Text) <> "" And Trim(txtNewLocation5.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation5.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo5.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo6.Text) <> "" And Trim(txtNewLocation6.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation6.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo6.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo7.Text) <> "" And Trim(txtNewLocation7.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation7.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo7.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo8.Text) <> "" And Trim(txtNewLocation8.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation8.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo8.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo9.Text) <> "" And Trim(txtNewLocation9.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation9.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo9.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo10.Text) <> "" And Trim(txtNewLocation10.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation10.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo10.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo11.Text) <> "" And Trim(txtNewLocation11.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation11.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo11.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo12.Text) <> "" And Trim(txtNewLocation12.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation12.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo12.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo13.Text) <> "" And Trim(txtNewLocation13.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation13.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo13.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo14.Text) <> "" And Trim(txtNewLocation14.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation14.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo14.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        If Trim(txtPartNo15.Text) <> "" And Trim(txtNewLocation15.Text) <> "" Then
            SQL_STATEMENT = "update PMIS_PARTMAS Set LOCATION = " & N2Str2Null(txtNewLocation15.Text) & " Where STOCKNO = " & N2Str2Null(txtPartNo15.Text)
            gconDMIS.Execute SQL_STATEMENT
        End If
        NEW_LogAudit "E", "UPDATE LOCATION", SQL_STATEMENT, "", "Parts", "", "", ""

        Call SetLocation

        Screen.MousePointer = 0
    End If

    Exit Sub
Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtNewLocation1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation1_LostFocus()
    If Trim(txtPartNo1.Text) <> "" And Trim(txtNewLocation1.Text) <> "" Then
        If CheckLocationAssigned(txtPartNo1.Text, txtNewLocation1.Text) <> "NA" Then
            If MsgBox("Location is already assigned to part number : " & CheckLocationAssigned(txtPartNo1.Text, txtNewLocation1.Text) & vbCrLf & " Accept duplication?", vbQuestion + vbYesNo, "Confirm...") = vbNo Then
                On Error Resume Next
                txtNewLocation1.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtNewLocation10_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation11_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation12_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation13_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation14_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation15_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation4_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation5_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation6_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation7_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation8_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNewLocation9_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO1_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO10_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO10_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO11_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO11_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO12_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO12_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO13_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO13_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO14_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO14_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO15_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO15_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO2_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO3_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO4_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO4_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO5_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO5_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO6_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO6_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO7_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO7_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO8_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO8_LostFocus()
    Call SetLocation
End Sub

Private Sub txtPARTNO9_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPARTNO9_LostFocus()
    Call SetLocation
End Sub

