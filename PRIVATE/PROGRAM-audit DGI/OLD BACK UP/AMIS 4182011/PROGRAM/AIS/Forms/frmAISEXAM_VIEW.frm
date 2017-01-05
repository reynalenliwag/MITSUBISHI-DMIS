VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmAISEXAM_VIEW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Exam Schedule"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAISEXAM_VIEW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8970
   Begin Crystal.CrystalReport rptDISPLAY 
      Left            =   1650
      Top             =   5490
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "List Of Applicants"
      Height          =   2955
      Left            =   60
      TabIndex        =   9
      Top             =   2490
      Width           =   8835
      Begin MSComctlLib.ListView lsvDISP 
         Height          =   2175
         Left            =   120
         TabIndex        =   0
         Top             =   330
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   3836
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Grade"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Remarks"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Note"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Double Click To Edit Grade"
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
         Height          =   165
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   2610
         Width           =   2475
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Schedule"
      Height          =   2385
      Left            =   60
      TabIndex        =   8
      Top             =   90
      Width           =   8805
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   4
         Left            =   1860
         TabIndex        =   7
         Top             =   1890
         Width           =   2115
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   1860
         TabIndex        =   6
         Top             =   1500
         Width           =   2115
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Description"
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   14
         Top             =   720
         Width           =   1680
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   1860
         TabIndex        =   4
         Top             =   660
         Width           =   6735
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   1860
         TabIndex        =   5
         Top             =   1080
         Width           =   3525
      End
      Begin VB.Label lblINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   1860
         TabIndex        =   3
         Top             =   240
         Width           =   2085
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   240
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Top             =   1950
         Width           =   255
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   240
         Index           =   0
         Left            =   1260
         TabIndex        =   12
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Exam"
         Height          =   240
         Index           =   2
         Left            =   330
         TabIndex        =   11
         Top             =   330
         Width           =   1380
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type Of Exam"
         Height          =   240
         Index           =   5
         Left            =   330
         TabIndex        =   10
         Top             =   1140
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdEXIT 
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
      Left            =   8100
      Picture         =   "frmAISEXAM_VIEW.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit Window"
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
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
      Left            =   7320
      Picture         =   "frmAISEXAM_VIEW.frx":0ADC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Exam Schedule"
      Top             =   5520
      Width           =   795
   End
   Begin VB.Label lblEXAMID 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   9000
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblSCHED_ID 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   9000
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "frmAISEXAM_VIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function DisplayList(SCHED_ID As Integer)
    Dim rsTmp                                                         As ADODB.Recordset
    Dim Item                                                          As ListItem

    lsvDISP.Enabled = False

    Set rsTmp = gconDMIS.Execute("Select * From HRMS_APPLICANT_EXAM_SCHEDULE Where Sched_ID = " & SCHED_ID & "")
    lsvDISP.ListItems.Clear

    If Not rsTmp.EOF And Not rsTmp.BOF Then
        lsvDISP.Enabled = True
    End If

    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = lsvDISP.ListItems.Add(, , rsTmp!APPLICANT_ID)
            Item.SubItems(1) = FindApplicantName(rsTmp!APPLICANT_ID)
            Item.SubItems(2) = Null2String(rsTmp!GRADE)
            Item.SubItems(3) = Null2String(rsTmp!ExamRemarks)
            Item.SubItems(4) = Null2String(rsTmp!NOTES)

            rsTmp.MoveNext
        Loop
    End If
End Function

Sub DisplayGrades()
    Dim rsTmp                                                         As ADODB.Recordset

    Set rsTmp = gconDMIS.Execute("Select Passing,MinScore,MaxScore From HRMS_EXAMTYPE Where ExamID = " & CLng(lblEXAMID.Caption) & "")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        frmAISEXAM_SAVE.lblINFO(6).Caption = rsTmp!MinScore
        frmAISEXAM_SAVE.lblINFO(7).Caption = rsTmp!Passing
        frmAISEXAM_SAVE.lblINFO(8).Caption = rsTmp!MaxScore
    End If
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
Private Sub cmdPrint_Click()
    Dim FILTER                                                        As String
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "ACESS_PRINT", "APPLICANT INFO") = False Then Exit Sub
    frmMain.MousePointer = 11
    rptDISPLAY.Formulas(0) = "DateOfExam = '" & lblINFO(0).Caption & "'"
    rptDISPLAY.Formulas(1) = "ExamDesc = '" & lblINFO(1).Caption & "'"
    rptDISPLAY.Formulas(2) = "ExamType = '" & lblINFO(2).Caption & "'"
    rptDISPLAY.Formulas(3) = "FTime = '" & lblINFO(3).Caption & "'"
    rptDISPLAY.Formulas(4) = "TTime = '" & lblINFO(4).Caption & "'"

    On Error GoTo ERROR

    FILTER = "{HRMS_EXAM_SCHEDULE.SCHED_ID} = " & CLng(lblSCHED_ID.Caption) & ""
    Call PrintSQLReport(rptDISPLAY, AIS_REPORT_PATH & "ExamSchedule.rpt", FILTER, AIS_REPORT_Connection, 1)

    frmMain.MousePointer = 0
    Exit Sub

ERROR:
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
    frmAISEXAM.Enabled = True
    On Error Resume Next
    frmAISEXAM.SetFocus
End Sub

Private Sub lsvDISP_DblClick()
    Dim INDEX                                                         As Long
    If Not lsvDISP.ListItems.count = 0 Then
        INDEX = lsvDISP.SelectedItem.INDEX
        With lsvDISP
            frmAISEXAM_VIEW.Enabled = False
            frmAISEXAM_SAVE.Show
            frmAISEXAM_SAVE.lblINFO(0).Caption = lblSCHED_ID.Caption
            frmAISEXAM_SAVE.lblINFO(1).Caption = .ListItems(INDEX).Text
            frmAISEXAM_SAVE.lblINFO(2).Caption = .ListItems(INDEX).SubItems(1)
            frmAISEXAM_SAVE.lblINFO(3).Caption = lblINFO(1).Caption
            frmAISEXAM_SAVE.lblINFO(4).Caption = lblINFO(0).Caption
            frmAISEXAM_SAVE.lblINFO(5).Caption = Trim(lblINFO(3).Caption & "-" & lblINFO(4).Caption)
            frmAISEXAM_SAVE.lblEXAMID.Caption = lblEXAMID.Caption

            Call DisplayGrades
            frmAISEXAM_SAVE.txtSCORE.Text = .ListItems(INDEX).SubItems(2)
            frmAISEXAM_SAVE.txtNOTE.Text = .ListItems(INDEX).SubItems(4)
            On Error Resume Next
            frmAISEXAM_SAVE.txtSCORE.SetFocus
        End With
    End If
End Sub

Private Sub lsvDISP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lsvDISP_DblClick
End Sub

