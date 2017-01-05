VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAISADD_EXAM 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2985
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5370
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5370
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
      Height          =   675
      Left            =   1440
      ScaleHeight     =   615
      ScaleWidth      =   3705
      TabIndex        =   9
      Top             =   2130
      Width           =   3765
      Begin VB.CommandButton cmdSADD_EXAMCANCEL 
         Caption         =   "CANCEL"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2490
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdADD_EXAMDELETE 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1260
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdADD_EXAMSAVE 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   30
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.TextBox txtEXAM_SCORE 
      Height          =   360
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1140
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPDateOfExam 
      Height          =   345
      Left            =   1920
      TabIndex        =   1
      Top             =   630
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      _Version        =   393216
      Format          =   53477377
      CurrentDate     =   39128
   End
   Begin VB.ComboBox cboEXAMType 
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3225
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   240
      Index           =   0
      Left            =   930
      TabIndex        =   11
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label Label1 
      Height          =   345
      Left            =   1950
      TabIndex        =   10
      Top             =   1590
      Width           =   1785
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Exam"
      Height          =   240
      Index           =   6
      Left            =   420
      TabIndex        =   8
      Top             =   720
      Width           =   1320
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      Height          =   240
      Index           =   7
      Left            =   1140
      TabIndex        =   7
      Top             =   1200
      Width           =   570
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exam Type"
      Height          =   240
      Index           =   20
      Left            =   660
      TabIndex        =   6
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmAISADD_EXAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSCHOOL_CANCEL_Click()
    Unload Me
End Sub

'Private Sub cmdADD_EXAMDELETE_Click()
'    Dim Sql              As String
'
'    If MsgBox("Are You Sure", vbQuestion + vbYesNo + vbDefaultButton2, "Delete This Exam") = vbYes Then
'        gconDMIS.Execute ("Delete From HRMS_APPLICANT_EXAM_SCHED Where ID = " & APPLICANT_ID & _
'                         " And Entry_ID = " & EXAM_ENTRY_ID & "")
'
'        Unload Me
'        Call frmAISEXAM_DISPLAY.DisplayApplicantInfoOnEXAMDISPLAY
'    Else
'        cboExamType.SetFocus
'    End If
'End Sub

Function CheckIfExamPass(EXAMID As Integer, ScoreGot As Double, lbl As Label) As String
    Dim rsTmp            As ADODB.Recordset
    Set rsTmp = GetRS("Select * From HRMS_ExamType Where ExamID = " & EXAMID)

    If Not (rsTmp.BOF And rsTmp.EOF) Then
        If ScoreGot >= rsTmp!Passing And ScoreGot < rsTmp!MaxScore Then
            CheckIfExamPass = rsTmp!PassRemark
            lbl.ForeColor = vbGreen

        ElseIf ScoreGot < rsTmp!Passing And ScoreGot >= rsTmp!MinScore Then

            CheckIfExamPass = rsTmp!MinRemark
            lbl.ForeColor = vbMagenta
        ElseIf ScoreGot >= rsTmp!MaxScore Then
            CheckIfExamPass = rsTmp!MaxRemark
            lbl.ForeColor = vbYellow
        Else
            CheckIfExamPass = "You Failed"
            lbl.ForeColor = vbRed
        End If

    End If
End Function

Function CheckIfExamAlrealyTaken(ID As Integer) As Boolean
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetRS("Select ExamType From HRMS_APPLICANT_EXAM_SCHED Where ExamType = " & ID & _
        " And Applicant_ID = " & APPLICANT_ID & "")
    
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        CheckIfExamAlrealyTaken = True
    Else
        CheckIfExamAlrealyTaken = False
    End If
    
    Set rsTmp = Nothing
End Function

'Private Sub cmdADD_EXAMSAVE_Click()
'    Dim VcboEXAMType As String, VcboEXAMStatus As String, VtxtEXAM_SCORE As String
'    Dim VDTPDateOfExam As String, Sql As String
'    Dim ID                  As Integer
'    Dim RESULT              As String
'    Dim ALREADY_TAKEN       As Boolean
'
'    VcboEXAMType = N2Str2Null(Right(cboEXAMType, 3))
'    VcboEXAMStatus = N2Str2Null(Label1.Caption)
'    VtxtEXAM_SCORE = N2Str2Null(txtEXAM_SCORE)
'    VDTPDateOfExam = N2Str2Null(DTPDateOfExam)
'
'    If Not cboEXAMType.Text = "" Then
'        If SAVE_OR_EDIT_EXAM = "SAVE" Then
'            ALREADY_TAKEN = CheckIfExamAlrealyTaken(Right(cboEXAMType, 3))
'            If ALREADY_TAKEN Then MsgBox "Exam Already Taken", vbInformation, "Examination": Exit Sub
'
'            Call GenerateNewID("HRMS_APPLICANT_EXAM_SCHED", ID)
'            EXAM_ENTRY_ID = ID
'
'            gconDMIS.Execute ("Insert Into HRMS_APPLICANT_EXAM_SCHED Values(" & APPLICANT_ID & _
'                              "," & EXAM_ENTRY_ID & _
'                              "," & VcboEXAMType & _
'                              "," & VDTPDateOfExam & _
'                              "," & VtxtEXAM_SCORE & _
'                              "," & VcboEXAMStatus & ")")
'
'            Unload Me
'            Call frmAISEXAM_DISPLAY.DisplayApplicantInfoOnEXAMDISPLAY
'        Else
'            gconDMIS.Execute "Update HRMS_APPLICANT_EXAM_SCHED Set ExamType = " & VcboEXAMType & _
'                              ",DateOfExam = " & VDTPDateOfExam & _
'                              ",Score = " & VtxtEXAM_SCORE & _
'                              ",Status = " & VcboEXAMStatus & _
'                            " Where Applicant_ID = " & APPLICANT_ID & " And Entry_ID = " & EXAM_ENTRY_ID
'
'            Unload Me
'            Call frmAISEXAM_DISPLAY.DisplayApplicantInfoOnEXAMDISPLAY
'        End If
'    Else
'        MsgBox "Incomplete Exam Information", vbExclamation, "Add Exam"
'        cboEXAMType.SetFocus
'    End If
'End Sub

Private Sub cmdSADD_EXAMCANCEL_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    Call FillTypeOfExamAndStatus
End Sub

Function FillTypeOfExamAndStatus()
    Dim Sql              As String
    Dim rsTmp            As New ADODB.Recordset
    Dim LENTOFID As String, SZERO As String

    Set rsTmp = gconDMIS.Execute("Select ID,ExamType From HRMS_ExamType Order By ExamType ASC")

    cboEXAMType.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            If Len(CStr(rsTmp!ID)) = 1 Then SZERO = "00"
            If Len(CStr(rsTmp!ID)) = 2 Then SZERO = "0"

            cboEXAMType.AddItem rsTmp!EXAMTYPE & " - " & SZERO & rsTmp!ID

            rsTmp.MoveNext
        Loop
    End If

    cboEXAMType.ListIndex = 0
    Set rsTmp = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    frmAISEXAM_DISPLAY.Enabled = True
    frmAISEXAM_DISPLAY.SetFocus
End Sub

Private Sub txtEXAM_SCORE_Change()
    If IsNumeric(txtEXAM_SCORE.Text) = True Then
        Label1.Caption = CheckIfExamPass(Right(cboEXAMType, 3), CDbl(txtEXAM_SCORE.Text), Label1)
    End If
End Sub

Private Sub txtEXAM_SCORE_Validate(Cancel As Boolean)
    If IsNumeric(txtEXAM_SCORE.Text) = False Then
        Cancel = True
    End If
End Sub
