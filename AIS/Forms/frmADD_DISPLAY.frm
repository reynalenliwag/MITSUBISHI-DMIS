VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmADD_EXAM 
   Caption         =   "Applicant Exam Taken"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCHILD_SAVE 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2970
      ScaleHeight     =   915
      ScaleWidth      =   5565
      TabIndex        =   7
      Top             =   4470
      Width           =   5625
      Begin VB.CommandButton cmdEXAM_PREV 
         Caption         =   "PREV"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   30
         Picture         =   "frmADD_DISPLAY.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   1365
      End
      Begin VB.CommandButton cmdEXAM_NEXT 
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   1410
         Picture         =   "frmADD_DISPLAY.frx":06EA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1365
      End
      Begin VB.CommandButton cmdEXAM_DISPLAY 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   2790
         Picture         =   "frmADD_DISPLAY.frx":0DD4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1365
      End
      Begin VB.CommandButton cmdEXAM_EXIT 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   4170
         Picture         =   "frmADD_DISPLAY.frx":0F9E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   1365
      End
   End
   Begin MSComctlLib.ListView lsvEXAM 
      Height          =   2865
      Left            =   90
      TabIndex        =   3
      Top             =   1440
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   5054
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblAPP 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   1770
      TabIndex        =   6
      Top             =   960
      Width           =   5115
   End
   Begin VB.Label lblAPP 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   1770
      TabIndex        =   5
      Top             =   540
      Width           =   5115
   End
   Begin VB.Label lblAPP 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   0
      Left            =   1770
      TabIndex        =   4
      Top             =   150
      Width           =   1815
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Applicant no."
      Height          =   240
      Index           =   20
      Left            =   270
      TabIndex        =   2
      Top             =   180
      Width           =   1305
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   240
      Index           =   7
      Left            =   540
      TabIndex        =   1
      Top             =   1020
      Width           =   1050
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last name"
      Height          =   240
      Index           =   6
      Left            =   540
      TabIndex        =   0
      Top             =   630
      Width           =   1020
   End
End
Attribute VB_Name = "frmADD_EXAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEXAM As New ADODB.Recordset

Private Sub cmdEXAM_EXIT_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Show
    Call CenterMe(mdiMAIN, Me, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmApplications.Enabled = True
    frmApplications.SetFocus
End Sub

Private Sub lsvEXAM_DblClick()
    If Not lsvEXAM.ListItems.Count = 0 Then
    
    
    
    End If
End Sub
