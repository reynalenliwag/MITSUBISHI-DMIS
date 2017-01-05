VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAISINTERVIEW_RESCHED 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reschedule Of Interview"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAISINTERVIEW_RESCHED.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   6540
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
      Height          =   795
      Left            =   5670
      Picture         =   "frmAISINTERVIEW_RESCHED.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exit Window"
      Top             =   6810
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "SCHEDULE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   90
      TabIndex        =   14
      Top             =   0
      Width           =   6300
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   1995
         TabIndex        =   25
         Top             =   255
         Width           =   4110
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Interview Description"
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
         Left            =   90
         TabIndex        =   24
         Top             =   360
         Width           =   1830
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
         Left            =   1485
         TabIndex        =   20
         Top             =   750
         Width           =   405
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
         Left            =   1470
         TabIndex        =   19
         Top             =   1125
         Width           =   435
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
         Left            =   4080
         TabIndex        =   18
         Top             =   1095
         Width           =   210
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   1995
         TabIndex        =   17
         Top             =   645
         Width           =   2565
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   1995
         TabIndex        =   16
         Top             =   1020
         Width           =   1605
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   4425
         TabIndex        =   15
         Top             =   1050
         Width           =   1605
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "NEW SCHEDULE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   60
      TabIndex        =   8
      Top             =   1635
      Width           =   6360
      Begin VB.ComboBox cboFTIME 
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
         Height          =   315
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1725
         Width           =   1635
      End
      Begin VB.ComboBox cboTTIME 
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
         Height          =   315
         Left            =   4470
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1725
         Width           =   1635
      End
      Begin VB.TextBox txtDESC 
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
         Height          =   840
         Left            =   1950
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   270
         Width           =   4200
      End
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
         Format          =   51904513
         CurrentDate     =   39148
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Interview Description"
         Height          =   240
         Index           =   9
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   2100
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
         TabIndex        =   12
         Top             =   1785
         Width           =   435
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
         TabIndex        =   11
         Top             =   1785
         Width           =   210
      End
      Begin VB.Label lblDATE 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Height          =   240
         Left            =   6540
         TabIndex        =   10
         Top             =   1650
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
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "LIST OF APPLICANT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   60
      TabIndex        =   7
      Top             =   3900
      Width           =   6360
      Begin MSComctlLib.ListView lsvLIST 
         Height          =   2010
         Left            =   90
         TabIndex        =   4
         Top             =   330
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   3545
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
      Left            =   4950
      Picture         =   "frmAISINTERVIEW_RESCHED.frx":0ADC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Save Schedule Of Interview"
      Top             =   6810
      Width           =   735
   End
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
      Height          =   1005
      Left            =   6690
      TabIndex        =   23
      Top             =   270
      Visible         =   0   'False
      Width           =   1635
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
      Left            =   1890
      TabIndex        =   22
      Top             =   7230
      Width           =   4935
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      Caption         =   "Conflict On Time:"
      Height          =   240
      Index           =   8
      Left            =   90
      TabIndex        =   21
      Top             =   6480
      Width           =   1725
   End
End
Attribute VB_Name = "frmAISINTERVIEW_RESCHED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CONFLICTS                                                         As Boolean
Dim RS_SCHED                                                          As ADODB.Recordset
Dim MR                                                                As Date

Function oPENsCHEDULE()
    Set RS_SCHED = gconDMIS.Execute("Select * From HRMS_INTERVIEW_SCHEDULE Where DateOfInterview = '" & frmAISINTERVIEW.dtpDATE1 & "' And INT_ID != " & CLng(lblSCHED_ID.Caption) & " Order By INT_ID ASC")
End Function

Function DisplayApplicantList()
    Dim rsTmp As ADODB.Recordset, rsSCHED As ADODB.Recordset, rsEDUC  As ADODB.Recordset
    Dim ITEM                                                          As ListItem
    Dim NOTES                                                         As String
    Dim INDEX                                                         As Integer
    Dim TIME1 As String, TIME2                                        As String
    Dim X_ID As Integer, t1 As Integer, T2                            As Integer

    INDEX = 1
    CONFLICTS = False

    lsvList.Enabled = False

    Set rsTmp = gconDMIS.Execute("Select * From HRMS_APPLICANT_INTERVIEW_SCHEDULE Where INT_ID = " & CLng(lblSCHED_ID.Caption) & "")

    lsvList.ListItems.Clear

    If Not rsTmp.EOF And rsTmp.BOF Then
        lsvList.Enabled = True
    End If

    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set rsSCHED = gconDMIS.Execute("Select * From HRMS_INTERVIEW_SCHEDULE Where DateOfInterview = '" & dtpDate & "'")
            If Not (rsSCHED.BOF And rsSCHED.EOF) Then
                Do While Not rsSCHED.EOF
                    t1 = rsSCHED!FROMTIME
                    T2 = rsSCHED!ToTime

                    If Not lblSCHED_ID.Caption = rsSCHED!INT_ID Then
                        '-----------------------------------------------------------------------------------------------------
                        If CLng(Right(cboFTIME, 3)) < 17 Then
                            If CLng(Right(cboFTIME, 3)) < rsSCHED!FROMTIME Then
                                If CLng(Right(cboTTIME, 3) - 1) >= rsSCHED!FROMTIME Then GoTo CONFLICT
                            End If

                            If CLng(Right(cboFTIME, 3)) > rsSCHED!FROMTIME Then
                                If CInt(Right(cboFTIME, 3)) <= rsSCHED!ToTime Then GoTo CONFLICT
                            End If

                            If CLng(Right(cboFTIME, 3)) = rsSCHED!FROMTIME Then GoTo CONFLICT
                        End If

                        '-----------------------------------------------------------------------------------------------------
                        '-----------------------------------------------------------------------------------------------------
                        If CLng(Right(cboFTIME, 3) - 1) > 16 Then
                            If CLng(Right(cboFTIME, 3) - 1) < rsSCHED!FROMTIME Then
                                If CLng(Right(cboTTIME, 3) - 2) >= rsSCHED!FROMTIME Then GoTo CONFLICT
                            End If

                            If CLng(Right(cboFTIME, 3) - 1) > rsSCHED!FROMTIME Then
                                If CInt(Right(cboFTIME, 3) - 1) <= rsSCHED!ToTime Then GoTo CONFLICT
                            End If

                            If CLng(Right(cboFTIME, 3) - 1) = rsSCHED!FROMTIME Then GoTo CONFLICT
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

                    NOTES = "CONFLICT SCHEDULE: "
                    NOTES = NOTES & " (" & TIME1 & " - " & TIME2 & ")"
                    CONFLICTS = True

JUMP1:
                    rsSCHED.MoveNext
                Loop
            End If
            Set ITEM = lsvList.ListItems.Add(, , rsTmp!APPLICANT_ID)
            ITEM.SubItems(1) = FindApplicantName(rsTmp!APPLICANT_ID)
            ITEM.SubItems(2) = NOTES
            lsvList.ListItems(INDEX).ListSubItems(2).ForeColor = vbRed

            NOTES = ""
            INDEX = INDEX + 1
            rsTmp.MoveNext
        Loop
    End If
End Function

Private Sub cboFTIME_Change()
    Dim rsTmp                                                         As ADODB.Recordset
    Dim SZERO                                                         As String

    If Right(cboFTIME, 2) < 17 Then Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME3 Where Time_ID > " & Right(cboFTIME, 2) & " And Time_ID < " & 18 & " Order By Time_ID ASC")
    If Right(cboFTIME, 2) > 16 Then Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME3 Where Time_ID > " & Right(cboFTIME, 2) & " Order By Time_ID ASC")

    cboTTIME.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            If Len(rsTmp!Time_ID) = 1 Then SZERO = "00"
            If Len(rsTmp!Time_ID) = 2 Then SZERO = "0"

            cboTTIME.AddItem rsTmp!Set_Time & Space(10) & SZERO & rsTmp!Time_ID

            rsTmp.MoveNext
        Loop
    End If
    cboTTIME.ListIndex = 0

    Call DisplayApplicantList
End Sub

Private Sub cboFTIME_Click()
    Call cboFTIME_Change
End Sub

Private Sub cboTTIME_Change()
    Call DisplayApplicantList
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:17
Private Sub cmdUPDATE_Click()
    Dim vtxtDESC As String, vdtpDATE                                  As String

    On Error GoTo Errorcode:

    If CDate(dtpDate) < Date Then
        MsgBox "Cannot Reschedule to a date that is already Passed"
        dtpDate.SetFocus
        Exit Sub
    End If

    frmMain.MousePointer = 11
    If Not CONFLICTS And lblNOTE.Caption = "" Then
        If Not txtDesc.Text = "" Then
            vtxtDESC = N2Str2Null(Trim(txtDesc.Text))
            vdtpDATE = dtpDate
            If MsgBox("Update Schedule", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
                If Right(cboFTIME, 3) < 17 Then
                    gconDMIS.Execute ("Update HRMS_INTERVIEW_SCHEDULE Set FromTime = " & CLng(Right(cboFTIME, 3)) & _
                                      ",ToTime = " & CLng(Right(cboTTIME, 3) - 1) & _
                                      ",DateOfInterview = '" & dtpDate & _
                                      "',InterviewDescription = " & vtxtDESC & " Where INT_ID = " & CLng(lblSCHED_ID.Caption) & "")
                End If
                If Right(cboFTIME, 3) > 16 Then
                    gconDMIS.Execute ("Update HRMS_INTERVIEW_SCHEDULE Set FromTime = " & CLng(Right(cboFTIME, 3) - 1) & _
                                      ",ToTime = " & CLng(Right(cboTTIME, 3) - 2) & _
                                      ",DateOfInterview = '" & vdtpDATE & _
                                      "',InterviewDescription = " & vtxtDESC & " Where INT_ID = " & CLng(lblSCHED_ID.Caption) & "")
                End If

                Unload Me
                Call frmAISINTERVIEW.FillSchedule1
                Call frmAISINTERVIEW.FillSchedule
            End If
        Else
            MsgBox "Enter Interview Description", vbExclamation, "Reschedule Of Exam"
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



    Exit Sub
    DATE1 = (Day(dtpDate)) + (MONTH(dtpDate)) + (YEAR(dtpDate))
    DATE2 = (Day(Date)) + (MONTH(Date)) + (YEAR(Date))
    If DATE1 >= DATE2 Then
        Call DisplayApplicantList
    Else
        MsgBox "You Cannot Reschedule Exam to a date already passed", vbCritical, "Reschedule of Exam"
        dtpDate.Day = 1
        dtpDate.YEAR = YEAR(lblDATE.Caption)
        dtpDate.MONTH = MONTH(lblDATE.Caption)
        dtpDate.Day = Day(lblDATE.Caption)
        On Error Resume Next
        dtpDate.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISINTERVIEW.Enabled = True
    On Error Resume Next
    frmAISINTERVIEW.SetFocus
End Sub

