VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAISEXAM_DISPLAY 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4500
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8775
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
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
   ScaleHeight     =   4500
   ScaleWidth      =   8775
   Begin VB.PictureBox picCHILD_SAVE 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   7740
      ScaleHeight     =   825
      ScaleWidth      =   855
      TabIndex        =   8
      Top             =   3660
      Width           =   855
      Begin VB.CommandButton cmdEXAM_EXIT 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   60
         Picture         =   "frmEXAM_DISPLAY.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   795
      End
   End
   Begin MSComctlLib.ListView lsvEXAM 
      Height          =   1995
      Left            =   90
      TabIndex        =   3
      Top             =   1620
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3519
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type of Exam"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date of Exam"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time Of Exam"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Score"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Remarks"
         Object.Width           =   3528
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8715
      _Version        =   655364
      _ExtentX        =   15372
      _ExtentY        =   450
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
   Begin VB.Label lblAPP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   1410
      TabIndex        =   2
      Top             =   1140
      Width           =   3645
   End
   Begin VB.Label lblAPP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   1410
      TabIndex        =   1
      Top             =   720
      Width           =   3645
   End
   Begin VB.Label lblAPP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   1410
      TabIndex        =   0
      Top             =   330
      Width           =   1305
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Applicant no."
      Height          =   240
      Index           =   20
      Left            =   90
      TabIndex        =   7
      Top             =   360
      Width           =   1305
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   240
      Index           =   7
      Left            =   240
      TabIndex        =   6
      Top             =   1140
      Width           =   1050
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   240
      Index           =   6
      Left            =   300
      TabIndex        =   5
      Top             =   750
      Width           =   1035
   End
End
Attribute VB_Name = "frmAISEXAM_DISPLAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function DisplayApplicantInfoOnEXAMDISPLAY()
    Dim SQL                                                           As String
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim rsSCHED                                                       As ADODB.Recordset
    Dim ITEM                                                          As ListItem
    Dim RESULT                                                        As String
    Dim TIME1 As String, TIME2                                        As String

    lsvEXAM.Enabled = False

    Set RSTMP = gconDMIS.Execute("Select FirstName,LastName From HRMS_APPLICANT_PERSONAL Where Applicant_ID = " & APPLICANT_ID & "")

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        frmAISEXAM_DISPLAY.lblAPP(0).Caption = APPLICANT_ID
        frmAISEXAM_DISPLAY.lblAPP(1).Caption = RSTMP!lastname
        frmAISEXAM_DISPLAY.lblAPP(2).Caption = RSTMP!FIRSTNAME
    End If

    SQL = "Select * From HRMS_APPLICANT_EXAM_SCHEDULE Where Applicant_ID = " & APPLICANT_ID & ""
    Set RSTMP = gconDMIS.Execute(SQL)

    lsvEXAM.ListItems.Clear

    If Not RSTMP.EOF And Not RSTMP.BOF Then
        lsvEXAM.Enabled = True
    End If

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set rsSCHED = gconDMIS.Execute("Select * From HRMS_EXAM_SCHEDULE Where SCHED_ID = " & RSTMP!SCHED_ID & "")
            If Not (rsSCHED.BOF And rsSCHED.EOF) Then
                RESULT = GetExamType(rsSCHED!EXAMID)

                Set ITEM = frmAISEXAM_DISPLAY.lsvEXAM.ListItems.Add(, , RESULT)
                ITEM.SubItems(1) = Null2String(rsSCHED!DATEofEXAM)
                If (rsSCHED!FROMTIME) < 9 Then
                    TIME1 = GetTime_TMP(rsSCHED!FROMTIME)
                    TIME2 = GetTime_TMP(rsSCHED!ToTime + 1)
                End If
                If (rsSCHED!FROMTIME) >= 9 Then
                    TIME1 = GetTime_TMP(rsSCHED!FROMTIME + 1)
                    TIME2 = GetTime_TMP(rsSCHED!ToTime + 2)
                End If
                ITEM.SubItems(2) = TIME1 & " - " & TIME2
                ITEM.SubItems(3) = Null2String(RSTMP!GRADE)
                ITEM.SubItems(4) = Null2String(RSTMP!ExamRemarks)
            End If

            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = Nothing
End Function

Private Sub cmdEXAM_EXIT_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISApplications.Enabled = True
    On Error Resume Next
    frmAISApplications.SetFocus
End Sub

