VERSION 5.00
Begin VB.Form frmAISMAIN2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Applicant Information System (AIS)"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMAIN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   8220
   Begin VB.CommandButton Command1 
      Height          =   795
      Left            =   4800
      Picture         =   "frmMAIN.frx":058A
      TabIndex        =   14
      ToolTipText     =   "Upload Applicants"
      Top             =   2880
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Timer tmeINTERVIEW 
      Interval        =   1000
      Left            =   1140
      Top             =   2760
   End
   Begin VB.CommandButton cmdAPP 
      Height          =   795
      Left            =   195
      Picture         =   "frmMAIN.frx":0DBC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Applicant Information Form"
      Top             =   75
      Width           =   885
   End
   Begin VB.CommandButton cmdSCHED 
      Height          =   795
      Left            =   210
      Picture         =   "frmMAIN.frx":15A0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Applicant Schedule of Interview "
      Top             =   2880
      Width           =   885
   End
   Begin VB.CommandButton cmdSEARCH 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   195
      Picture         =   "frmMAIN.frx":1E21
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Search for an Applicants"
      Top             =   1950
      Width           =   885
   End
   Begin VB.CommandButton cmdUPLOAD 
      Height          =   795
      Left            =   4800
      Picture         =   "frmMAIN.frx":27F8
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Upload Applicants"
      Top             =   1980
      Width           =   915
   End
   Begin VB.CommandButton cmdEXAMTYPE 
      Height          =   795
      Left            =   4800
      Picture         =   "frmMAIN.frx":302A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exam Types"
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdPOSIION 
      Height          =   795
      Left            =   4800
      Picture         =   "frmMAIN.frx":392F
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Open Position"
      Top             =   1035
      Width           =   915
   End
   Begin VB.Timer tmeEXAMINATION 
      Interval        =   1000
      Left            =   1110
      Top             =   870
   End
   Begin VB.CommandButton cmdEXAM 
      Height          =   795
      Left            =   180
      Picture         =   "frmMAIN.frx":40EF
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Schedule of Exam"
      Top             =   1005
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   6
      Left            =   5850
      TabIndex        =   15
      Top             =   3150
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Position"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   5
      Left            =   5850
      TabIndex        =   13
      Top             =   1260
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Exams"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   4
      Left            =   5850
      TabIndex        =   12
      ToolTipText     =   "Type of Exams"
      Top             =   390
      Width           =   2115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Application Form"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   3
      Left            =   1140
      TabIndex        =   11
      Top             =   330
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   2
      Left            =   5850
      TabIndex        =   10
      Top             =   2190
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule Of Interview"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   1
      Left            =   1170
      TabIndex        =   9
      Top             =   3090
      Width           =   3195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inquiry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1170
      TabIndex        =   8
      Top             =   2145
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule Of Exam"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   0
      Left            =   1140
      TabIndex        =   7
      Top             =   1275
      Width           =   2625
   End
End
Attribute VB_Name = "frmAISMAIN2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAPP_Click()
    If ApplySecurityValidation = True Then
        If Module_Access(LOGID, "AIS MASTER APPLICATION", "DATA ENTRY") = False Then Exit Sub
    End If

    frmAISApplications.Show
    frmAISApplications.tbcApplication.SelectedItem = 0
End Sub

Private Sub cmdEXAM_Click()
    frmAISEXAM.Show
End Sub

Private Sub cmdEXAMTYPE_Click()
    If Module_Access(LOGID, "APPLICANT EXAM TYPE", "DATA ENTRY") = False Then Exit Sub
    frmAISAdd_TYPEofEXAM.Show
End Sub

Private Sub cmdPOSIION_Click()
    If Module_Access(LOGID, "APPLICANT OPEN POSITION", "DATA ENTRY") = False Then Exit Sub
    frmAISPOSITION.Show
End Sub

Private Sub cmdSCHED_Click()
    frmAISINTERVIEW.Show
    'frmAIS_SCHEDULE.Show
End Sub

Private Sub cmdSEARCH_Click()
    If Module_Access(LOGID, "APPLICANT INQUIRY", "INQUIRY") = False Then Exit Sub
    frmAISSEARCH.Show
End Sub

Private Sub cmdUPLOAD_Click()
    If Module_Access(LOGID, "UPLOAD APPLICANT", "PROCESSING") = False Then Exit Sub
    On Error GoTo Errorcode:

    frmAISPOSITION_APPLY.Show
    On Error Resume Next
    frmAISPOSITION_APPLY.txtSearch.SetFocus





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Command1_Click()
    frmPASmain.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Debug.Print KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

