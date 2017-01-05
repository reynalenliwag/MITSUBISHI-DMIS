VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H8000000D&
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Splash.frx":000C
   ScaleHeight     =   2880
   ScaleWidth      =   5895
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1980
      ScaleHeight     =   555
      ScaleWidth      =   3195
      TabIndex        =   19
      Top             =   810
      Width           =   3195
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Office Supplies Management System"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   60
         TabIndex        =   20
         Top             =   30
         Width           =   3105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   30
      ScaleHeight     =   285
      ScaleWidth      =   5805
      TabIndex        =   1
      Top             =   2250
      Width           =   5835
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   390
         TabIndex        =   16
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   750
         TabIndex        =   15
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   1110
         TabIndex        =   14
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   1470
         TabIndex        =   13
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   1830
         TabIndex        =   12
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   2190
         TabIndex        =   11
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   2550
         TabIndex        =   10
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   8
         Left            =   2910
         TabIndex        =   9
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   9
         Left            =   3270
         TabIndex        =   8
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   10
         Left            =   3630
         TabIndex        =   7
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   11
         Left            =   3990
         TabIndex        =   6
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   12
         Left            =   4350
         TabIndex        =   5
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   13
         Left            =   4710
         TabIndex        =   4
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   14
         Left            =   5070
         TabIndex        =   3
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lab 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   15
         Left            =   5430
         TabIndex        =   2
         Top             =   30
         Width           =   315
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   135
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label labCon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting to Access Database... Please wait..."
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   2610
      Width           =   5745
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Gray, WBlue As String
Dim WBlueCount, CNT As Integer

Private Sub Command1_Click()
Unload Me
On Error Resume Next
Unload frmSecurity
End Sub

Private Sub Form_Load()
Gray = &HE0E0E0: WBlue = &HE37331
WBlueCount = 0
CenterMe frmMain, Me, 1
End Sub

Private Sub Timer1_Timer()
WBlueCount = WBlueCount + 1
If WBlueCount > 15 Then WBlueCount = 0
For CNT = 0 To 15
    lab(CNT).BackColor = Gray
Next
lab(WBlueCount).BackColor = WBlue
End Sub


