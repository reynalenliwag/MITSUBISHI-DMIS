VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAISEXAM_RESCHED 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reschedule Of Exam"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAISEXAM_RESCHED.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   6750
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
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
      Left            =   5790
      Picture         =   "frmAISEXAM_RESCHED.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exit Window"
      Top             =   6540
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
      Height          =   765
      Left            =   5010
      Picture         =   "frmAISEXAM_RESCHED.frx":141C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Update Schedule Of Exam"
      Top             =   6540
      Width           =   795
   End
   Begin VB.Frame Frame3 
      Caption         =   "LIST OF APPLICANT"
      Height          =   2610
      Left            =   90
      TabIndex        =   9
      Top             =   3840
      Width           =   6510
      Begin MSComctlLib.ListView lsvLIST 
         Height          =   2130
         Left            =   90
         TabIndex        =   4
         Top             =   330
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   3757
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Notes"
            Object.Width           =   9701
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "NEW SCHEDULE"
      Height          =   2205
      Left            =   90
      TabIndex        =   8
      Top             =   1590
      Width           =   6510
      Begin MSComCtl2.DTPicker dtpDATE 
         Height          =   375
         Left            =   1935
         TabIndex        =   1
         Top             =   1200
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   49545217
         CurrentDate     =   39148
      End
      Begin VB.TextBox txtDESC 
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
         Height          =   840
         Left            =   1935
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   270
         Width           =   4410
      End
      Begin VB.ComboBox cboTTIME 
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
         Height          =   315
         Left            =   4470
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1725
         Width           =   1635
      End
      Begin VB.ComboBox cboFTIME 
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
         Height          =   315
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1725
         Width           =   1635
      End
      Begin VB.Label lblDATE 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Height          =   240
         Left            =   6630
         TabIndex        =   25
         Top             =   1470
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblFID 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7170
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Exam Description"
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
         Index           =   6
         Left            =   180
         TabIndex        =   16
         Top             =   285
         Width           =   1500
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   4110
         TabIndex        =   15
         Top             =   1785
         Width           =   210
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Index           =   4
         Left            =   1395
         TabIndex        =   14
         Top             =   1785
         Width           =   435
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Date"
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
         Index           =   3
         Left            =   1395
         TabIndex        =   13
         Top             =   1335
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SCHEDULE"
      Height          =   1500
      Left            =   90
      TabIndex        =   7
      Top             =   30
      Width           =   6510
      Begin VB.Label lblSCHED_ID 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   90
         TabIndex        =   26
         Top             =   750
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   4365
         TabIndex        =   21
         Top             =   1020
         Width           =   1605
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   1785
         TabIndex        =   20
         Top             =   1020
         Width           =   1605
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   1785
         TabIndex        =   19
         Top             =   645
         Width           =   2565
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Exam Description"
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
         Left            =   180
         TabIndex        =   18
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   1785
         TabIndex        =   17
         Top             =   255
         Width           =   4560
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   3810
         TabIndex        =   12
         Top             =   1095
         Width           =   210
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Left            =   1260
         TabIndex        =   11
         Top             =   1065
         Width           =   435
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Date"
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
         Left            =   1275
         TabIndex        =   10
         Top             =   720
         Width           =   405
      End
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      Caption         =   "Conflict On Time:"
      Height          =   240
      Index           =   8
      Left            =   90
      TabIndex        =   24
      Top             =   6570
      Width           =   1725
   End
   Begin VB.Label lblNOTE 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   1920
      TabIndex        =   23
      Top             =   6600
      Width           =   4935
   End
End
Attribute VB_Name = "frmAISEXAM_RESCHED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CONFLICTS                                                         As Boolean
Dim RS_SCHED                                                          As ADODB.Recordset
Dim MR                                                                As Date

Function oPENsCHEDULE()
    Set RS_SCHED = gconDMIS.Execute("Select * From HRMS_EXAM_SCHEDULE Where DateOfExam = '" & frmAISEXAM.dtpDATE1 & "' And Sched_ID != " & CLng(lblSCHED_ID.Caption) & " And ExamID = " & Right(frmAISEXAM.cboEXAMTYPE1, 3) & "  Order By Sched_ID ASC")
End Function

Function DisplayApplicantList()
    Dim rsTmp As ADODB.Recordset, rsSCHED As ADODB.Recordset, rsEDUC  As ADODB.Recordset
    Dim Item                                                          As ListItem
    Dim NOTES                                                         As String
    Dim INDEX                                                         As Integer
    Dim TIME1 As String, TIME2                                        As String
    Dim X_ID As Integer, t1 As Integer, T2                            As Integer

    lsvLIST.Enabled = False

    INDEX = 1
    CONFLICTS = False

    Set rsTmp = gconDMIS.Execute("Select * From HRMS_APPLICANT_EXAM_SCHEDULE Where SCHED_ID = " & CLng(lblSCHED_ID.Caption) & "")

    lsvLIST.ListItems.Clear

    If Not rsTmp.EOF And Not rsTmp.BOF Then
        lsvLIST.Enabled = True
    End If



    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set rsEDUC = gconDMIS.Execute("Select * From HRMS_APPLICANT_EXAM_SCHEDULE Where Applicant_ID = " & rsTmp!APPLICANT_ID & " Order By Sched_ID ASC")

            If Not (rsEDUC.BOF And rsEDUC.EOF) Then
                Do While Not rsEDUC.EOF
                    If Not CLng(lblSCHED_ID.Caption) = rsEDUC!SCHED_ID Then
                        Set rsSCHED = gconDMIS.Execute("Select * From HRMS_EXAM_SCHEDULE Where SCHED_ID = " & CLng(rsEDUC!SCHED_ID) & "")

                        If Not (rsSCHED.BOF And rsSCHED.EOF) Then
                            X_ID = rsSCHED!EXAMID
                            t1 = rsSCHED!FROMTIME
                            T2 = rsSCHED!ToTime

                            '-----------------------------------------------------------------------------------------------------
                            'If CLng(Right(cboFTIME, 3)) < 9 Then
                            If CLng(Right(cboFTIME, 3)) < 17 Then
                                If CLng(Right(cboFTIME, 3)) < rsSCHED!FROMTIME Then
                                    If CLng(Right(cboTTIME, 3) - 1) >= rsSCHED!FROMTIME Then
                                        If rsSCHED!DATEofEXAM = dtpDATE Then
                                            GoTo CONFLICT
                                        End If
                                    End If
                                End If

                                If CLng(Right(cboFTIME, 3)) > rsSCHED!FROMTIME Then
                                    If CInt(Right(cboFTIME, 3)) <= rsSCHED!ToTime Then
                                        If rsSCHED!DATEofEXAM = dtpDATE Then
                                            GoTo CONFLICT
                                        End If
                                    End If
                                End If

                                If CLng(Right(cboFTIME, 3)) = rsSCHED!FROMTIME Then
                                    If rsSCHED!DATEofEXAM = dtpDATE Then
                                        GoTo CONFLICT
                                    End If
                                End If
                            End If
                            '-----------------------------------------------------------------------------------------------------

                            '-----------------------------------------------------------------------------------------------------
                            If CLng(Right(cboFTIME, 3) - 1) > 16 Then
                                If CLng(Right(cboFTIME, 3) - 1) < rsSCHED!FROMTIME Then
                                    'If CLng(Right(cboTTIME, 3) - 1) >= rsSCHED!FROMTIME Then
                                    If CLng(Right(cboTTIME, 3) - 2) >= rsSCHED!FROMTIME Then
                                        If rsSCHED!DATEofEXAM = dtpDATE Then
                                            GoTo CONFLICT
                                        End If
                                    End If
                                End If

                                If CLng(Right(cboFTIME, 3) - 1) > rsSCHED!FROMTIME Then
                                    If CInt(Right(cboFTIME, 3) - 1) <= rsSCHED!ToTime Then
                                        If rsSCHED!DATEofEXAM = dtpDATE Then
                                            GoTo CONFLICT
                                        End If
                                    End If
                                End If

                                If CLng(Right(cboFTIME, 3) - 1) = rsSCHED!FROMTIME Then
                                    If rsSCHED!DATEofEXAM = dtpDATE Then
                                        GoTo CONFLICT
                                    End If
                                End If
                            End If
                            '-----------------------------------------------------------------------------------------------------

                            GoTo JUMP1

CONFLICT:
                            'If (T1) < 9 Then
                            If (t1) < 17 Then
                                TIME1 = GetTime_TMP(t1)
                                TIME2 = GetTime_TMP(T2 + 1)
                            End If
                            'If (T1) >= 9 Then
                            If (t1) >= 17 Then
                                TIME1 = GetTime_TMP(t1 + 1)
                                TIME2 = GetTime_TMP(T2 + 2)
                            End If

                            NOTES = "CONFLICT SCHEDULE: "
                            NOTES = NOTES & ReturnExamType(rsSCHED!EXAMID) & " (" & TIME1 & " - " & TIME2 & ")"
                            CONFLICTS = True

JUMP1:
                        End If
                    End If
                    rsEDUC.MoveNext
                Loop
            End If

            Set Item = lsvLIST.ListItems.Add(, , rsTmp!APPLICANT_ID)
            Item.SubItems(1) = FindApplicantName(rsTmp!APPLICANT_ID)
            Item.SubItems(2) = NOTES
            lsvLIST.ListItems(INDEX).ListSubItems(2).ForeColor = vbRed

            NOTES = ""
            INDEX = INDEX + 1
            rsTmp.MoveNext
        Loop
    End If
End Function

Sub CheckForDuplicateWithOtherSchedule()
    Dim NOTES                                                         As String
    Dim TIME1 As String, TIME2                                        As String
    Dim X_ID As Integer, t1 As Integer, T2                            As Integer

    If Not (RS_SCHED.BOF And RS_SCHED.EOF) Then
        RS_SCHED.MoveFirst
        Do While Not RS_SCHED.EOF
            X_ID = RS_SCHED!EXAMID
            t1 = RS_SCHED!FROMTIME
            T2 = RS_SCHED!ToTime

            If CLng(Right(cboFTIME, 3)) < 17 Then
                If CLng(Right(cboFTIME, 3)) < RS_SCHED!FROMTIME Then
                    If CLng(Right(cboTTIME, 3) - 1) >= RS_SCHED!FROMTIME Then
                        If RS_SCHED!DATEofEXAM = dtpDATE Then GoTo CONFLICT
                    End If
                End If

                If CLng(Right(cboFTIME, 3)) > RS_SCHED!FROMTIME Then
                    If CInt(Right(cboFTIME, 3)) <= RS_SCHED!ToTime Then
                        If RS_SCHED!DATEofEXAM = dtpDATE Then GoTo CONFLICT
                    End If
                End If

                If CLng(Right(cboFTIME, 3)) = RS_SCHED!FROMTIME Then
                    If RS_SCHED!DATEofEXAM = dtpDATE Then GoTo CONFLICT
                End If
            End If
            '-----------------------------------------------------------------------------------------------------

            '-----------------------------------------------------------------------------------------------------
            If CLng(Right(cboFTIME, 3) - 1) > 16 Then
                If CLng(Right(cboFTIME, 3) - 1) < RS_SCHED!FROMTIME Then
                    If CLng(Right(cboTTIME, 3) - 2) >= RS_SCHED!FROMTIME Then
                        If RS_SCHED!DATEofEXAM = dtpDATE Then GoTo CONFLICT
                    End If
                End If

                If CLng(Right(cboFTIME, 3) - 1) > RS_SCHED!FROMTIME Then
                    If CInt(Right(cboFTIME, 3) - 1) <= RS_SCHED!ToTime Then
                        If RS_SCHED!DATEofEXAM = dtpDATE Then GoTo CONFLICT
                    End If
                End If

                If CLng(Right(cboFTIME, 3) - 1) = RS_SCHED!FROMTIME Then
                    If RS_SCHED!DATEofEXAM = dtpDATE Then GoTo CONFLICT
                End If
            End If
            '-----------------------------------------------------------------------------------------------------

            GoTo JUMP1

CONFLICT:
            If (t1) < 17 Then
                TIME1 = GetTime_TMP(t1)
                TIME2 = GetTime_TMP(T2 + 1)
            End If
            If (t1) >= 17 Then
                TIME1 = GetTime_TMP(t1 + 1)
                TIME2 = GetTime_TMP(T2 + 2)
            End If

            'NOTES = "CONFLICT SCHEDULE: "
            NOTES = ReturnExamType(RS_SCHED!EXAMID) & " (" & TIME1 & " - " & TIME2 & ")"

            GoTo NOTING

JUMP1:

            RS_SCHED.MoveNext
        Loop
    End If

    '    Exit Sub

NOTING:
    lblNOTE.Caption = NOTES

End Sub

Private Sub cboFTIME_Change()
    Dim rsTmp                                                         As ADODB.Recordset
    Dim SZERO                                                         As String

    'If Right(cboFTIME, 2) < 9 Then Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME1 Where Time_ID > " & Right(cboFTIME, 2) & " And Time_ID < " & 10 & " Order By Time_ID ASC")
    'If Right(cboFTIME, 2) > 8 Then Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME1 Where Time_ID > " & Right(cboFTIME, 2) & " Order By Time_ID ASC")

    If Right(cboFTIME, 2) < 17 Then Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME3 Where Time_ID > " & Right(cboFTIME, 2) & " And Time_ID < " & 18 & " Order By Time_ID ASC")
    If Right(cboFTIME, 2) > 16 Then Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME3 Where Time_ID > " & Right(cboFTIME, 2) & " Order By Time_ID ASC")

    cboTTIME.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            If Len(rsTmp!Time_ID) = 1 Then SZERO = "00"
            If Len(rsTmp!Time_ID) = 2 Then SZERO = "0"

            'cboTTIME.AddItem rsTmp!SETTIME & Space(10) & SZERO & rsTmp!Time_ID
            cboTTIME.AddItem rsTmp!Set_Time & Space(10) & SZERO & rsTmp!Time_ID

            rsTmp.MoveNext
        Loop
    End If
    cboTTIME.ListIndex = 0

    Call DisplayApplicantList
    Call CheckForDuplicateWithOtherSchedule
End Sub

Private Sub cboFTIME_Click()
    Call cboFTIME_Change
End Sub

Private Sub cboTTIME_Change()
    Call DisplayApplicantList
    Call CheckForDuplicateWithOtherSchedule
End Sub

Private Sub cboTTIME_Click()
    Call cboTTIME_Change
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
    Dim vtxtDESC As String, vdtpDATE                                  As String

    On Error GoTo Errorcode:

    If Not CONFLICTS And lblNOTE.Caption = "" Then
        If Not txtDesc.Text = "" Then
            vtxtDESC = N2Str2Null(txtDesc.Text)
            vdtpDATE = dtpDATE
            If MsgBox("Update Schedule", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
                frmMain.MousePointer = 11
                'If Right(cboFTIME, 3) < 9 Then
                If Right(cboFTIME, 3) < 17 Then
                    gconDMIS.Execute ("Update HRMS_EXAM_SCHEDULE Set FromTime = " & CLng(Right(cboFTIME, 3)) & _
                                      ",ToTime = " & CLng(Right(cboTTIME, 3) - 1) & _
                                      ",DateOfExam = '" & dtpDATE & _
                                      "',ActivityDescription = " & vtxtDESC & " Where Sched_ID = " & CLng(lblSCHED_ID.Caption) & "")
                End If
                'If Right(cboFTIME, 3) > 8 Then
                If Right(cboFTIME, 3) > 16 Then
                    gconDMIS.Execute ("Update HRMS_EXAM_SCHEDULE Set FromTime = " & CLng(Right(cboFTIME, 3) - 1) & _
                                      ",ToTime = " & CLng(Right(cboTTIME, 3) - 2) & _
                                      ",DateOfExam = '" & vdtpDATE & _
                                      "',ActivityDescription = " & vtxtDESC & " Where Sched_ID = " & CLng(lblSCHED_ID.Caption) & "")
                End If

                Unload Me
                Call frmAISEXAM.FillSchedule1
                Call frmAISEXAM.FillSchedule
            End If
        Else
            MsgBox "Enter Exam Description", vbExclamation, "Reschedule Of Exam"
            On Error Resume Next
            txtDesc.SetFocus
        End If
    Else
        MsgBox "Theres a Conflict on Schedule", vbExclamation, "Reschedule of Exam"
        On Error Resume Next
        cboFTIME.SetFocus
    End If

    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub dtpDATE_Change()
    Dim DATE1                                                         As Long
    Dim DATE2                                                         As Long

    DATE1 = (Day(dtpDATE)) + (Month(dtpDATE)) + (YEAR(dtpDATE))
    DATE2 = (Day(Date)) + (Month(Date)) + (YEAR(Date))
    If DATE1 >= DATE2 Then
        Call DisplayApplicantList
        Call CheckForDuplicateWithOtherSchedule
    Else
        MsgBox "You Cannot Reschedule Exam to a date already passed", vbCritical, "Reschedule of Exam"
        dtpDATE.Day = 1
        dtpDATE.YEAR = YEAR(lblDATE.Caption)
        dtpDATE.Month = Month(lblDATE.Caption)
        dtpDATE.Day = Day(lblDATE.Caption)
        On Error Resume Next
        dtpDATE.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISEXAM.Enabled = True
    On Error Resume Next
    frmAISEXAM.SetFocus
End Sub

