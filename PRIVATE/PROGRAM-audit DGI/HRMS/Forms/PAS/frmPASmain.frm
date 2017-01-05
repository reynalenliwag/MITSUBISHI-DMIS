VERSION 5.00
Begin VB.Form frmPASmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personnel Appraisal System"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPASmain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   7845
   Begin VB.CommandButton frmFORMD 
      Caption         =   "FORM D"
      Height          =   765
      Left            =   3780
      TabIndex        =   3
      Top             =   1380
      Width           =   915
   End
   Begin VB.CommandButton frmFORMC 
      Caption         =   "FORM C"
      Height          =   765
      Left            =   3780
      TabIndex        =   2
      Top             =   210
      Width           =   915
   End
   Begin VB.CommandButton cmdFORMB 
      Caption         =   "FORM B"
      Height          =   765
      Left            =   300
      TabIndex        =   1
      Top             =   1380
      Width           =   915
   End
   Begin VB.CommandButton cmdFORMA 
      Caption         =   "FORM A"
      Height          =   765
      Left            =   300
      TabIndex        =   0
      Top             =   210
      Width           =   915
   End
   Begin VB.Label lblCAP 
      Caption         =   "OVERALL PERFORMANCE SUMMARY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   555
      Index           =   3
      Left            =   4920
      TabIndex        =   7
      Top             =   1590
      Width           =   2865
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      Caption         =   "BEHAVIOR FACTORS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   6
      Top             =   510
      Width           =   2070
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      Caption         =   "SKILLS FACTOR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   1
      Left            =   1350
      TabIndex        =   5
      Top             =   1650
      Width           =   1590
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      Caption         =   "FORM A"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   0
      Left            =   1350
      TabIndex        =   4
      Top             =   510
      Width           =   795
   End
End
Attribute VB_Name = "frmPASmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFORMA_Click()
    frmPASformA.Show
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

    Call CenterMe(frmMain, Me, 1)
End Sub

