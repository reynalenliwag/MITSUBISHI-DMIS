VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAISSchedule_EDIT 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4620
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8985
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
   ScaleHeight     =   4620
   ScaleWidth      =   8985
   Begin VB.PictureBox picMENU 
      Height          =   615
      Left            =   5340
      ScaleHeight     =   555
      ScaleWidth      =   3495
      TabIndex        =   7
      Top             =   3930
      Width           =   3555
      Begin VB.CommandButton cmdPREV 
         Caption         =   "BACK"
         Height          =   555
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1155
      End
      Begin VB.CommandButton cmdNEXT 
         Caption         =   "NEXT"
         Height          =   555
         Left            =   1170
         TabIndex        =   5
         Top             =   0
         Width           =   1155
      End
      Begin VB.CommandButton cmdEXIT 
         Caption         =   "EXIT"
         Height          =   555
         Left            =   2340
         TabIndex        =   6
         Top             =   0
         Width           =   1155
      End
   End
   Begin MSComctlLib.ListView lsvAPPLIST 
      Height          =   2415
      Left            =   150
      TabIndex        =   3
      Top             =   1410
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   4260
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Position"
         Object.Width           =   7408
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date of Exam"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time of Exam"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblAPP 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   990
      Width           =   6945
   End
   Begin VB.Label lblAPP 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   570
      Width           =   6945
   End
   Begin VB.Label lblAPP 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   180
      Width           =   1815
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Applicant no."
      Height          =   240
      Index           =   20
      Left            =   420
      TabIndex        =   10
      Top             =   240
      Width           =   1305
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   240
      Index           =   7
      Left            =   690
      TabIndex        =   9
      Top             =   1080
      Width           =   1050
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last name"
      Height          =   240
      Index           =   6
      Left            =   690
      TabIndex        =   8
      Top             =   660
      Width           =   1020
   End
End
Attribute VB_Name = "frmAISSchedule_EDIT"
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
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    frmAISSchedule.Enabled = True
    frmAISSchedule.SetFocus
End Sub

Private Sub lsvAPPLIST_Click()
    If Not lsvAPPLIST.ListItems.Count = 0 Then
    
        With lsvAPPLIST
        
        End With
    End If
End Sub

