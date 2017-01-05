VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAISEXAM_SET 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule of Examinee"
   ClientHeight    =   7590
   ClientLeft      =   750
   ClientTop       =   1920
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAISSET_EXAM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   7305
   Begin VB.Frame FmeLIST 
      Enabled         =   0   'False
      Height          =   4665
      Left            =   90
      TabIndex        =   14
      Top             =   2880
      Width           =   7125
      Begin VB.CommandButton cmdCANCEL 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   6150
         Picture         =   "frmAISSET_EXAM.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exit Window"
         Top             =   3780
         Width           =   795
      End
      Begin VB.TextBox txtSEARCH 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1050
         TabIndex        =   4
         Top             =   270
         Width           =   4125
      End
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5370
         Picture         =   "frmAISSET_EXAM.frx":0ADC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save Entry"
         Top             =   3780
         Width           =   795
      End
      Begin MSComctlLib.ListView LsvAPP 
         Height          =   2295
         Left            =   150
         TabIndex        =   5
         Top             =   1110
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4048
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "FullName"
            Object.Width           =   6174
         EndProperty
      End
      Begin MSComctlLib.ListView lsvLIST 
         Height          =   2295
         Left            =   3690
         TabIndex        =   8
         Top             =   1110
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   4048
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Name"
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - SEARCH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5340
         TabIndex        =   23
         Top             =   450
         Width           =   1200
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Double Click to Remove"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   8
         Left            =   3690
         TabIndex        =   22
         Top             =   3480
         Width           =   2235
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Double Click to Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   21
         Top             =   3510
         Width           =   1860
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LIST OF EXAMINEE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   3690
         TabIndex        =   18
         Top             =   840
         Width           =   1800
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LIST OF APPLICANT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   17
         Top             =   840
         Width           =   1890
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame FmeSCHED 
      Height          =   2925
      Left            =   90
      TabIndex        =   11
      Top             =   -60
      Width           =   7125
      Begin VB.ComboBox cboTIMEofEXAM 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2310
         Width           =   1635
      End
      Begin VB.CommandButton cmdCAN 
         Caption         =   "CANCEL RESET"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5490
         Picture         =   "frmAISSET_EXAM.frx":117C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancel/Reset Exam"
         Top             =   1950
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdSET 
         Caption         =   "SET"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4530
         Picture         =   "frmAISSET_EXAM.frx":1808
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Set Schedule"
         Top             =   1950
         Width           =   975
      End
      Begin VB.TextBox txtDESC 
         Appearance      =   0  'Flat
         Height          =   900
         Left            =   2040
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblSCHED_ID 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   735
         Left            =   6630
         TabIndex        =   24
         Top             =   1410
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Type"
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
         Index           =   5
         Left            =   960
         TabIndex        =   20
         Top             =   1410
         Width           =   960
      End
      Begin VB.Label lblExamType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   1260
         Width           =   4215
      End
      Begin VB.Label lblFTIME 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   285
         Left            =   0
         TabIndex        =   19
         Top             =   1830
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activity Description"
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
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label lblFROMTIME 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   1740
         Width           =   1605
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Time"
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
         Index           =   0
         Left            =   960
         TabIndex        =   13
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Time"
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
         Index           =   7
         Left            =   1380
         TabIndex        =   12
         Top             =   2400
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmAISEXAM_SET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RECALL                                                            As String
Dim TITLE                                                             As String

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:17
Private Sub cmdCAN_Click()
    On Error GoTo Errorcode:

    txtDesc.Text = TITLE
    cboTIMEofEXAM.Text = RECALL
    cmdSET.Caption = "CHANGE"
    cmdCAN.Visible = False
    txtDesc.Enabled = False
    cboTIMEofEXAM.Enabled = False
    FmeLIST.Enabled = True
    On Error Resume Next
    txtSEARCH.SetFocus
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:17
Private Sub cmdSave_Click()
    Dim EXAM_DESC As String, DATEofEXAM                               As String
    Dim GRADE As String, REMARKS As String, NOTE                      As String
    Dim rsTmp                                                         As ADODB.Recordset

    On Error GoTo Errorcode:

    frmMain.MousePointer = 11
    If Not lsvList.ListItems.count = 0 Then
        If MsgBox("Save This Exam Schedule", vbQuestion + vbYesNo + vbDefaultButton1, "Are You Sure") = vbYes Then
            If txtDesc.Text = "" Then
                MsgBox "Enter a Activity Description", vbExclamation, "Schedule Of Exam"
                On Error Resume Next
                txtDesc.SetFocus
                Exit Sub
            End If

            EXAM_DESC = N2Str2Null(Trim(txtDesc.Text))
            DATEofEXAM = N2Str2Null(frmAISEXAM.dtpDate)
            GRADE = N2Str2Null("")
            REMARKS = N2Str2Null("")
            NOTE = N2Str2Null("")

            gconDMIS.Execute ("Insert Into HRMS_EXAM_SCHEDULE Values(" & CInt(lblSCHED_ID) & _
                              "," & EXAM_DESC & "," & CInt(Right(frmAISEXAM.cboExamType, 3)) & "," & DATEofEXAM & _
                              "," & CInt(lblFTIME) & "," & CInt(Right(cboTIMEofEXAM, 2)) & ")")

            Set rsTmp = gconDMIS.Execute("Select * From HRMS_APPLICANT_EXAM_SCHEDULE_TMP Order By Applicant_ID ASC")
            If Not (rsTmp.BOF And rsTmp.EOF) Then
                Do While Not rsTmp.EOF
                    gconDMIS.Execute ("Insert Into HRMS_APPLICANT_EXAM_SCHEDULE Values(" & rsTmp!APPLICANT_ID & _
                                      "," & CLng(lblSCHED_ID) & "," & GRADE & "," & REMARKS & "," & NOTE & ")")

                    rsTmp.MoveNext
                Loop
            End If

            Unload Me
            Call frmAISEXAM.FillSchedule
            Call frmAISEXAM.FillSchedule1
        Else

            If LsvAPP.ListItems.count > 0 And LsvAPP.Enabled = True Then: LsvAPP.SetFocus
        End If
    End If

    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub cmdSET_Click()
    If cmdSET.Caption = "RESET" Or cmdSET.Caption = "SET" Then
        If Not txtDesc.Text = "" Then
            If Not lsvList.ListItems.count = 0 Then
                If MsgBox("Changing the (to)Time of the Exam can Affect or Conflict the Schedule of the Applicant, COntinue", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
                    Call CheckForConflict
                    Call DisplayChange
                    GoTo JUMP1
                End If
            Else
JUMP1:
                cboTIMEofEXAM.Enabled = False
                txtDesc.Enabled = False
                cmdCAN.Visible = False
                cmdSET.Caption = "CHANGE"
                FmeLIST.Enabled = True
                On Error Resume Next
                txtSEARCH.SetFocus
            End If
        Else
            MsgBox "Activity Descripion cannot be Blank", vbInformation, "Schedule of Examination"
            On Error Resume Next
            txtDesc.SetFocus
            Exit Sub
        End If
    Else                                                      'CHANGE
        TITLE = Trim(txtDesc.Text)
        RECALL = cboTIMEofEXAM.Text
        cmdSET.Caption = "RESET"

        txtDesc.Enabled = True
        cmdCAN.Visible = True
        cboTIMEofEXAM.Enabled = True
        FmeLIST.Enabled = False

        txtDesc.SetFocus
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        If FmeLIST.Enabled = True Then
            txtSEARCH.Text = ""
            On Error Resume Next
            txtSEARCH.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    Call txtsearch_Change
    Call GenerateNewSCHED_ID
End Sub

Private Sub GenerateNewSCHED_ID()
    Dim rsTmp                                                         As ADODB.Recordset

    lblSCHED_ID.Caption = 0
    Set rsTmp = gconDMIS.Execute("Select TOP 1  SCHED_ID FROM HRMS_EXAM_SCHEDULE Order By SCHED_ID DESC")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        lblSCHED_ID.Caption = rsTmp!SCHED_ID
    End If
    lblSCHED_ID.Caption = lblSCHED_ID.Caption + 1

    gconDMIS.Execute ("Delete From HRMS_APPLICANT_EXAM_SCHEDULE_TMP")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISEXAM.Enabled = True
    On Error Resume Next
    frmAISEXAM.SetFocus
End Sub

Private Sub LsvAPP_DblClick()
    Dim INDEX                                                         As Long

    If Not LsvAPP.ListItems.count = 0 Then
        INDEX = LsvAPP.SelectedItem.INDEX
        With LsvAPP
            Call CheckConflictDuplicate(CInt(.ListItems(INDEX).Text), INDEX)
        End With
    End If
End Sub

Private Sub DisplayChange()
    Dim rsTmp                                                         As ADODB.Recordset
    Dim ITEM                                                          As ListItem

    lsvList.Enabled = False
    Set rsTmp = gconDMIS.Execute("Select * From HRMS_APPLICANT_EXAM_SCHEDULE_TMP")
    lsvList.ListItems.Clear

    If Not rsTmp.EOF And rsTmp.BOF Then
        lsvList.Enabled = True
    End If

    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvList.ListItems.Add(, , rsTmp!APPLICANT_ID)
            ITEM.SubItems(1) = rsTmp!FULLNAME

            rsTmp.MoveNext
        Loop

    End If
End Sub

Private Sub CheckForConflict()
    Dim rsTmp As ADODB.Recordset, rsSCHED                             As ADODB.Recordset

    Set rsTmp = gconDMIS.Execute("Select * From HRMS_APPLICANT_EXAM_SCHEDULE ORder By Applicant_ID")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set rsSCHED = gconDMIS.Execute("Select * From HRMS_EXAM_SCHEDULE Where SCHED_ID = " & CLng(rsTmp!SCHED_ID) & "")
            If Not (rsSCHED.BOF And rsSCHED.EOF) Then

                '-----------------------------------------------------------------------------------------------------
                If CInt(lblFTIME.Caption) < rsSCHED!FROMTIME Then
                    If CInt(Right(cboTIMEofEXAM, 2)) >= rsSCHED!FROMTIME Then
                        gconDMIS.Execute ("Delete From HRMS_APPLICANT_EXAM_SCHEDULE_TMP Where Applicant_ID = " & rsTmp!APPLICANT_ID & "")
                    End If
                End If
                If CInt(lblFTIME.Caption) = rsSCHED!FROMTIME Then
                    gconDMIS.Execute ("Delete From HRMS_APPLICANT_EXAM_SCHEDULE_TMP Where Applicant_ID = " & rsTmp!APPLICANT_ID & "")
                End If
                If CInt(lblFTIME.Caption) > rsSCHED!FROMTIME Then
                    If CInt(lblFTIME.Caption) <= rsSCHED!ToTime Then
                        gconDMIS.Execute ("Delete From HRMS_APPLICANT_EXAM_SCHEDULE_TMP Where Applicant_ID = " & rsTmp!APPLICANT_ID & "")
                    End If
                End If
                '-----------------------------------------------------------------------------------------------------

            End If

            rsTmp.MoveNext
        Loop
    End If
End Sub

Private Sub CheckConflictDuplicate(APP_ID As Integer, INDEX As Long)
    Dim rsTmp As ADODB.Recordset, rsExam As ADODB.Recordset, rsSCHED  As ADODB.Recordset

    Dim ITEM                                                          As ListItem
    Dim X_ID                                                          As Integer
    Dim NOTES As String, TIME1 As String, TIME2                       As String
    Dim t1 As Integer, T2                                             As Integer

    Set rsTmp = gconDMIS.Execute("Select SCHED_ID,EXAMREMARKS From HRMS_APPLICANT_EXAM_SCHEDULE Where APPLICANT_ID = " & APP_ID & "")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set rsSCHED = gconDMIS.Execute("Select * From HRMS_EXAM_SCHEDULE Where SCHED_ID = " & rsTmp!SCHED_ID & "")

            If Not (rsSCHED.BOF And rsSCHED.EOF) Then
                X_ID = rsSCHED!EXAMID
                t1 = rsSCHED!FROMTIME
                T2 = rsSCHED!ToTime

                If rsSCHED!EXAMID = CInt(Right(frmAISEXAM.cboExamType, 3)) Then
                    If Null2String(rsTmp!ExamRemarks) = "" Then    'EXAMREMARKS -> NO RESULT
                        MsgBox "Applicant Already Scheduled to Take This Exam", vbInformation, "Schedule of Exam"
                        If LsvAPP.ListItems.count > 0 And LsvAPP.Enabled = True Then: LsvAPP.SetFocus

                        GoTo ALREADY_TAKE
                    ElseIf Null2String(rsTmp!ExamRemarks) = "Passed" Then    'EXAMREMARKS -> PASSED
                        MsgBox "Applicant Already Passed The Exam", vbInformation, "Schedule Of Exam"

                        If LsvAPP.ListItems.count > 0 And LsvAPP.Enabled = True Then: LsvAPP.SetFocus
                        GoTo ALREADY_TAKE
                    Else                                      'EXAMREMARKS -> FAILED or EXAM NOT YET TAKEN
                        GoTo JUMP1
                    End If
                Else                                          'CHECK FOR CONFLICTS
JUMP1:
                    '-----------------------------------------------------------------------------------------------------
                    If rsSCHED!DATEofEXAM = frmAISEXAM.dtpDate Then
                        If CInt(lblFTIME.Caption) < rsSCHED!FROMTIME Then
                            If CInt(Right(cboTIMEofEXAM, 2)) >= rsSCHED!FROMTIME Then
                                GoTo CONFLICT
                            End If
                        End If
                        If CInt(lblFTIME.Caption) > rsSCHED!FROMTIME Then
                            If CInt(lblFTIME.Caption) <= rsSCHED!ToTime Then
                                GoTo CONFLICT
                            End If
                        End If
                        If CInt(lblFTIME.Caption) = rsSCHED!FROMTIME Then
                            GoTo CONFLICT
                        End If
                    End If
                End If
            End If
            rsTmp.MoveNext
        Loop
        GoTo JUMP_ELSE
    Else
JUMP_ELSE:
        Set rsExam = gconDMIS.Execute("Select * From HRMS_APPLICANT_EXAM_SCHEDULE_TMP Where APPLICANT_ID = " & _
                                      APP_ID & "")
        If (rsExam.BOF And rsExam.EOF) Then
            gconDMIS.Execute ("Insert Into HRMS_APPLICANT_EXAM_SCHEDULE_TMP Values(" & APP_ID & _
                              ",'" & LsvAPP.ListItems(INDEX).SubItems(1) & "')")

            Set ITEM = lsvList.ListItems.Add(, , APP_ID)
            ITEM.SubItems(1) = LsvAPP.ListItems(INDEX).SubItems(1)
        Else
            'ON THE LIST ALREADY......
        End If
    End If

    Exit Sub

ALREADY_TAKE:

    Exit Sub

CONFLICT:
    If (t1) < 9 Then
        TIME1 = GetTime_TMP(t1)
        TIME2 = GetTime_TMP(T2 + 1)
    End If
    If (t1) >= 9 Then
        TIME1 = GetTime_TMP(t1 + 1)
        TIME2 = GetTime_TMP(T2 + 2)
    End If

    NOTES = "CONFLICT SCHEDULE: "
    NOTES = NOTES & ReturnExamType(X_ID) & " (" & TIME1 & " - " & TIME2 & ")"
    MsgBox NOTES, vbExclamation, "Schedule Of Exam"
    LsvAPP.SetFocus
End Sub

Private Sub LsvAPP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call LsvAPP_DblClick
End Sub

Private Sub lsvLIST_DblClick()
    Dim INDEX                                                         As Integer
    If Not lsvList.ListItems.count = 0 Then
        INDEX = lsvList.SelectedItem.INDEX
        With lsvList
            If MsgBox("Remove This Applicant From the List", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
                gconDMIS.Execute ("Delete From HRMS_APPLICANT_EXAM_SCHEDULE_TMP Where Applicant_ID = " & CLng(.ListItems(INDEX).Text) & "")
                Call DisplayChange
            End If
        End With
    End If
End Sub

Private Sub txtsearch_Change()
    Dim Keyword                                                       As String
    Dim rsTmp                                                         As ADODB.Recordset
    Dim ITEM                                                          As ListItem

    LsvAPP.Enabled = False

    Keyword = Trim(txtSEARCH)
    Set rsTmp = gconDMIS.Execute("Select Applicant_ID,FirstName,Lastname From HRMS_APPLICANT_PERSONAL Where FirstName Like '%" & _
                                 Keyword & "%' Or Lastname Like '%" & Keyword & "%' Order by Applicant_ID ASC")

    LsvAPP.ListItems.Clear

    If Not rsTmp.EOF And Not rsTmp.BOF Then
        LsvAPP.Enabled = True
    End If

    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = LsvAPP.ListItems.Add(, , rsTmp!APPLICANT_ID)
            ITEM.SubItems(1) = Null2String(rsTmp!lastname & ", " & rsTmp!FIRSTNAME)

            rsTmp.MoveNext
        Loop
    End If

End Sub

