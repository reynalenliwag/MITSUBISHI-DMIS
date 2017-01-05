VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAISEXAM_EDIT 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4410
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8760
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8760
   Begin MSComctlLib.ListView lsvEXAMLIST 
      Height          =   2145
      Left            =   120
      TabIndex        =   3
      Top             =   1380
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3784
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Exam Type "
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date of Exam"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time of Exam"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Score Got"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox picMENU 
      Height          =   615
      Left            =   5280
      ScaleHeight     =   555
      ScaleWidth      =   3345
      TabIndex        =   10
      Top             =   3630
      Width           =   3405
      Begin VB.CommandButton cmdNEXT 
         Caption         =   "NEXT"
         Height          =   555
         Left            =   1140
         TabIndex        =   5
         Top             =   0
         Width           =   1125
      End
      Begin VB.CommandButton cmdPREV 
         Caption         =   "PREVIOUS"
         Height          =   555
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1125
      End
      Begin VB.CommandButton cmdEXIT 
         Caption         =   "EXIT"
         Height          =   555
         Left            =   2190
         TabIndex        =   6
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   240
      Index           =   6
      Left            =   630
      TabIndex        =   9
      Top             =   630
      Width           =   1035
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   240
      Index           =   7
      Left            =   630
      TabIndex        =   8
      Top             =   1050
      Width           =   1050
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Applicant no."
      Height          =   240
      Index           =   20
      Left            =   360
      TabIndex        =   7
      Top             =   210
      Width           =   1305
   End
   Begin VB.Label lblAPP 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   0
      Left            =   1860
      TabIndex        =   0
      Top             =   150
      Width           =   1815
   End
   Begin VB.Label lblAPP 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   1860
      TabIndex        =   1
      Top             =   540
      Width           =   6795
   End
   Begin VB.Label lblAPP 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   1860
      TabIndex        =   2
      Top             =   960
      Width           =   6795
   End
End
Attribute VB_Name = "frmAISEXAM_EDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAPP As ADODB.Recordset

Private Sub cmdEXIT_Click()
    Unload Me
End Sub

Private Sub cmdNEXT_Click()
    rsAPP.MoveNext
    If rsAPP.EOF Then
        rsAPP.MoveLast
    End If
        
    Call DisplayApplicantInformation
End Sub

Private Sub cmdPREV_Click()
    rsAPP.MovePrevious
    If rsAPP.BOF Then
        rsAPP.MoveFirst
    End If
    Call DisplayApplicantInformation
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    Call OpenAppRecord

    rsAPP.MoveFirst
    Call DisplayApplicantInformation
End Sub

Private Sub OpenAppRecord()
    Set rsAPP = New ADODB.Recordset
    rsAPP.Open "Select * From HRMS_APPLICANT_PERSONAL Order By LastName ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Function DisplayApplicantInformation()
    Dim rsTmp As ADODB.Recordset, rsExam As ADODB.Recordset, rsSCHED As ADODB.Recordset
    Dim rsEXAMTYPE As ADODB.Recordset
    Dim ITEM As ListItem
    Dim ExamDescription As String, SZERO As String
    
    Set rsTmp = GetRS("Select FirstName,LastName,Applicant_ID From HRMS_APPLICANT_PERSONAL Where Applicant_ID = " & _
            rsAPP!APPLICANT_ID & "")
        
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        lblAPP(0).Caption = Null2String(rsTmp!APPLICANT_ID)
        lblAPP(1).Caption = Null2String(rsTmp!LastName)
        lblAPP(2).Caption = Null2String(rsTmp!FirstName)
        
        Set rsExam = GetRS("Select * From HRMS_APPLICANT_EXAM_SCHED Where Applicant_ID = " & rsAPP!APPLICANT_ID & " Order By Exam_Sched_ID ASC")
        
        lsvEXAMLIST.ListItems.Clear
        If Not (rsExam.BOF And rsExam.EOF) Then
            Do While Not rsExam.EOF
                Set rsSCHED = GetRS("Select * From HRMS_EXAM_SCHEDULE WHERE EXAM_SCHED_ID = " & rsExam!EXAM_SCHED_ID & "")
                
                If Not (rsSCHED.BOF And rsSCHED.EOF) Then
                    Set ITEM = lsvEXAMLIST.ListItems.Add(, , rsExam!EXAM_SCHED_ID)
                    ExamDescription = GetExamType(rsSCHED!EXAMTYPE)
                    If Len(rsSCHED!EXAMTYPE) = 1 Then SZERO = "00"
                    If Len(rsSCHED!EXAMTYPE) = 2 Then SZERO = "0"
                    
                    ITEM.SubItems(1) = ExamDescription & " - " & SZERO & rsSCHED!EXAMTYPE
                    ITEM.SubItems(2) = rsSCHED!DATEofEXAM
                    ITEM.SubItems(3) = rsSCHED!TimeOfExam
                    ITEM.SubItems(4) = rsExam!Score
                    ITEM.SubItems(5) = rsExam!Status
                End If
                
                rsExam.MoveNext
            Loop
        End If
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    'frmMain.tbMENU.Enabled = False
'    frmAISEXAM_SCHED.Enabled = True
'    frmAISEXAM_SCHED.SetFocus
End Sub

Private Sub lsvEXAMLIST_DblClick()
    Dim INDEX As Integer
    Dim rsTmp As ADODB.Recordset
    
    If Not lsvEXAMLIST.ListItems.Count = 0 Then
        INDEX = CInt(lsvEXAMLIST.SelectedItem.INDEX)
        With lsvEXAMLIST
            frmAISEXAM_EDIT.Enabled = False
            frmAISEXAM_SAVE.Show
            
            frmAISEXAM_SAVE.lblINFO(0).Caption = .ListItems(INDEX).Text
            frmAISEXAM_SAVE.lblINFO(1).Caption = lblAPP(0).Caption
            frmAISEXAM_SAVE.lblINFO(2).Caption = lblAPP(1).Caption & "," & lblAPP(2).Caption
            frmAISEXAM_SAVE.lblINFO(3).Caption = .ListItems(INDEX).SubItems(1)
            frmAISEXAM_SAVE.lblINFO(4).Caption = .ListItems(INDEX).SubItems(2)
            frmAISEXAM_SAVE.lblINFO(5).Caption = .ListItems(INDEX).SubItems(3)
            frmAISEXAM_SAVE.txtSCORE.Text = .ListItems(INDEX).SubItems(4)
            
            Set rsTmp = GetRS("Select Passing,MaxScore,MinScore From HRMS_ExamType Where ExamID = " & _
                CInt(Right(.ListItems(INDEX).SubItems(1), 3)) & "")
            If Not (rsTmp.BOF And rsTmp.EOF) Then
                frmAISEXAM_SAVE.lblINFO(6).Caption = rsTmp!MinScore
                frmAISEXAM_SAVE.lblINFO(7).Caption = rsTmp!Passing
                frmAISEXAM_SAVE.lblINFO(8).Caption = rsTmp!MaxScore
            End If
        End With
    End If
End Sub
