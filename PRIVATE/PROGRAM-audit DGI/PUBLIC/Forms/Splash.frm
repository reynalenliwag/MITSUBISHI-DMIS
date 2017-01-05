VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4455
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   330
      Top             =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      ScaleHeight     =   285
      ScaleWidth      =   5745
      TabIndex        =   1
      Top             =   3810
      Width           =   5775
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
         Width           =   285
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   135
      Left            =   5910
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label labCompName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Looking for Authentication..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   2460
      TabIndex        =   19
      Top             =   1980
      Width           =   3345
   End
   Begin VB.Image Image1 
      Height          =   3795
      Left            =   0
      Picture         =   "Splash.frx":000C
      Top             =   0
      Width           =   5925
   End
   Begin VB.Label labCon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting to Access Database... Please wait..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   18
      Top             =   4170
      Width           =   5625
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Gray                                                    As String
Dim WBlue                                                   As String
Dim WBlueCount                                              As Integer
Dim cnt                                                     As Integer

Private Sub Command1_Click()
    Unload Me
    On Error Resume Next

End Sub

Private Sub Form_Load()
    Gray = &HE0E0E0: WBlue = &HE37331
    WBlueCount = 0:
    If COMPANY_NAME <> "" Then labCompName.Caption = Cap1st(COMPANY_NAME): DoEvents
    CenterMe Screen, Me, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'frmMain.ApplyPatches
End Sub

Private Sub Timer1_Timer()
    If COMPANY_NAME <> "" Then labCompName.Caption = Cap1st(COMPANY_NAME): DoEvents
    WBlueCount = WBlueCount + 1
    If WBlueCount > 15 Then WBlueCount = 0
    For cnt = 0 To 15
        lab(cnt).BackColor = Gray
    Next
    lab(WBlueCount).BackColor = WBlue
End Sub

