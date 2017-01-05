VERSION 5.00
Begin VB.Form frmAISEXAM_SAVE 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6045
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7665
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7665
   Begin VB.Frame Frame1 
      Caption         =   "Grades"
      Height          =   3165
      Left            =   4020
      TabIndex        =   14
      Top             =   1770
      Width           =   3525
      Begin VB.TextBox txtNOTE 
         Appearance      =   0  'Flat
         Height          =   900
         Left            =   930
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   2130
         Width           =   2475
      End
      Begin VB.TextBox txtSCORE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1890
         TabIndex        =   0
         Top             =   1650
         Width           =   1485
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   240
         Index           =   7
         Left            =   210
         TabIndex        =   29
         Top             =   2130
         Width           =   570
      End
      Begin VB.Label lblINFO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   1890
         TabIndex        =   13
         Top             =   1230
         Width           =   1485
      End
      Begin VB.Label lblINFO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   1890
         TabIndex        =   12
         Top             =   750
         Width           =   1485
      End
      Begin VB.Label lblINFO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   1890
         TabIndex        =   11
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score Got"
         Height          =   240
         Index           =   3
         Left            =   720
         TabIndex        =   26
         Top             =   1740
         Width           =   990
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Grade"
         Height          =   240
         Index           =   9
         Left            =   180
         TabIndex        =   17
         Top             =   1320
         Width           =   1530
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Grade"
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   16
         Top             =   390
         Width           =   1470
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passing Grade"
         Height          =   240
         Index           =   4
         Left            =   330
         TabIndex        =   15
         Top             =   840
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   5100
      ScaleHeight     =   825
      ScaleWidth      =   2475
      TabIndex        =   28
      Top             =   5040
      Width           =   2475
      Begin VB.CommandButton cmdEXIT 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1620
         Picture         =   "frmEXAM_SAVE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCANCEL 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   840
         Picture         =   "frmEXAM_SAVE.frx":0552
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Delete Entry"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUPDATE 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   60
         Picture         =   "frmEXAM_SAVE.frx":0C3C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Save Entry"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Applicant Information"
      Height          =   1725
      Left            =   150
      TabIndex        =   22
      Top             =   60
      Width           =   6195
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   1860
         TabIndex        =   7
         Top             =   1200
         Width           =   4185
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   1860
         TabIndex        =   6
         Top             =   780
         Width           =   1665
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1860
         TabIndex        =   5
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   240
         Index           =   0
         Left            =   750
         TabIndex        =   25
         Top             =   1260
         Width           =   945
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule Code"
         Height          =   240
         Index           =   20
         Left            =   240
         TabIndex        =   24
         Top             =   450
         Width           =   1470
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Applicant No."
         Height          =   240
         Index           =   6
         Left            =   420
         TabIndex        =   23
         Top             =   840
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Grades"
      Height          =   1785
      Left            =   150
      TabIndex        =   18
      Top             =   1770
      Width           =   3795
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   1590
         TabIndex        =   10
         Top             =   1320
         Width           =   2025
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   1590
         TabIndex        =   9
         Top             =   810
         Width           =   2025
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   1590
         TabIndex        =   8
         Top             =   360
         Width           =   2025
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Of Exam"
         Height          =   240
         Index           =   5
         Left            =   150
         TabIndex        =   21
         Top             =   1380
         Width           =   1350
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type Of Exam"
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   20
         Top             =   420
         Width           =   1380
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Exam"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   19
         Top             =   900
         Width           =   1350
      End
   End
   Begin VB.Label lblEXAMID 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   180
      TabIndex        =   30
      Top             =   4200
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   150
      TabIndex        =   27
      Top             =   3630
      Visible         =   0   'False
      Width           =   3795
   End
End
Attribute VB_Name = "frmAISEXAM_SAVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function CheckIfExamPass(EXAMID As Integer, ScoreGot As Double, lbl As Label) As String
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_ExamType Where ExamID = " & EXAMID)

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If ScoreGot >= RSTMP!Passing And ScoreGot < RSTMP!MaxScore Then
            CheckIfExamPass = RSTMP!PassRemark
            lbl.ForeColor = vbGreen
        ElseIf ScoreGot < RSTMP!Passing And ScoreGot >= RSTMP!MinScore Then
            CheckIfExamPass = RSTMP!MinRemark
            lbl.ForeColor = vbMagenta
        ElseIf ScoreGot >= RSTMP!MaxScore Then
            CheckIfExamPass = RSTMP!MaxRemark
            lbl.ForeColor = vbYellow
        Else
            CheckIfExamPass = "Failed"
            lbl.ForeColor = vbRed
        End If

    End If
End Function

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:16
Private Sub cmdCancel_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "ACESS_DELETE", "APPLICANT INFO") = False Then Exit Sub
    If MsgBox("Cancel this Exam", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
        gconDMIS.Execute ("Delete From HRMS_APPLICANT_EXAM_SCHEDULE Where APPLICANT_ID = " & _
                          CInt(lblINFO(1)) & " And Sched_ID = " & CInt(lblINFO(0)) & "")

        Call frmAISEXAM_VIEW.DisplayList(CLng(lblINFO(0)))
        Unload Me
    End If

    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:16
Private Sub cmdUPDATE_Click()
    Dim vtxtNOTES                                                     As String

    On Error GoTo Errorcode:

    If IsNumeric(txtSCORE) = True Then
        frmMain.MousePointer = 11
        If MsgBox("Update Exam result", vbQuestion + vbYesNo + vbDefaultButton1, "Are You Sure") = vbYes Then
            vtxtNOTES = N2Str2Null(txtNOTE)

            gconDMIS.Execute ("Update HRMS_APPLICANT_EXAM_SCHEDULE Set ExamRemarks = " & N2Str2Null(Label1) & _
                              ",Grade = " & CDbl(txtSCORE) & _
                              ",Notes = " & vtxtNOTES & _
                            " Where Applicant_ID = " & CInt(lblINFO(1)) & _
                            " And Sched_ID = " & CInt(lblINFO(0)) & "")

            Call frmAISEXAM_VIEW.DisplayList(CLng(lblINFO(0)))
            Unload Me
        End If
    Else
        MsgBox "Enter a Valid Score Got", vbExclamation, "Edit Exam"
        On Error Resume Next
        txtSCORE.SetFocus
    End If

    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISEXAM_VIEW.Enabled = True
    On Error Resume Next
    frmAISEXAM_VIEW.SetFocus
End Sub

Private Sub txtSCORE_Change()
    If IsNumeric(txtSCORE.Text) = True Then
        Label1.Caption = CheckIfExamPass(CLng(lblEXAMID.Caption), CDbl(txtSCORE.Text), Label1)
    End If
End Sub

