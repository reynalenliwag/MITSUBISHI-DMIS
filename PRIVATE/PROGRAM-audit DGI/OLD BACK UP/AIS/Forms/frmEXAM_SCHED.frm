VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAISEXAM_SCHED 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule of Examinee"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
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
   ScaleHeight     =   7770
   ScaleWidth      =   11130
   Begin VB.Frame Frame1 
      Caption         =   "Schedule Examinee of Given Time"
      Height          =   3975
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   4845
      Begin MSComctlLib.ListView lsvSCHED 
         Height          =   3165
         Left            =   120
         TabIndex        =   9
         Top             =   390
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   5583
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ExamType"
            Object.Width           =   2646
         EndProperty
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
         Index           =   2
         Left            =   150
         TabIndex        =   24
         Top             =   3660
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmdEDIT_EXAM 
      Caption         =   "&UPDATE EXAM"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   7170
      Width           =   1695
   End
   Begin VB.CommandButton cmdEXIT 
      Caption         =   "&EXIT"
      Height          =   495
      Left            =   3210
      TabIndex        =   11
      Top             =   7170
      Width           =   1695
   End
   Begin VB.Frame fmeMAIN 
      Caption         =   "Schedule of Exam"
      Height          =   2745
      Left            =   120
      TabIndex        =   17
      Top             =   90
      Width           =   4815
      Begin VB.ComboBox cboKINDofEXAM 
         Height          =   360
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   3015
      End
      Begin VB.ComboBox cboTIMEofEXAM 
         Height          =   360
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton cmdSETTIME 
         Caption         =   "SET"
         Height          =   495
         Left            =   2970
         TabIndex        =   3
         Top             =   2040
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPDATEofEXAM 
         Height          =   345
         Left            =   1650
         TabIndex        =   1
         Top             =   930
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   609
         _Version        =   393216
         Format          =   50593793
         CurrentDate     =   39129
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Exam"
         Height          =   240
         Index           =   6
         Left            =   180
         TabIndex        =   20
         Top             =   1020
         Width           =   1320
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Of Exam"
         Height          =   240
         Index           =   7
         Left            =   150
         TabIndex        =   19
         Top             =   1530
         Width           =   1350
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kind Of Exam"
         Height          =   240
         Index           =   20
         Left            =   210
         TabIndex        =   18
         Top             =   540
         Width           =   1290
      End
   End
   Begin VB.Frame fmeSCHED 
      Caption         =   "Applicant Schedule to Take the Exam"
      Enabled         =   0   'False
      Height          =   3195
      Left            =   5070
      TabIndex        =   16
      Top             =   4470
      Width           =   5955
      Begin VB.CommandButton cmdCANCEL 
         Caption         =   "CANCEL SCHEDULE"
         Height          =   585
         Left            =   4290
         TabIndex        =   13
         Top             =   2460
         Width           =   1545
      End
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "SAVE SCHEDULE"
         Height          =   585
         Left            =   2640
         TabIndex        =   12
         Top             =   2460
         Width           =   1575
      End
      Begin MSComctlLib.ListView lsvAPPEXAM 
         Height          =   1995
         Left            =   120
         TabIndex        =   8
         Top             =   390
         Width           =   5715
         _ExtentX        =   10081
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Name"
            Object.Width           =   6174
         EndProperty
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
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   2550
         Width           =   2235
      End
   End
   Begin VB.Frame fmeAPP 
      Caption         =   "Choose Applicant"
      Enabled         =   0   'False
      Height          =   4245
      Left            =   5070
      TabIndex        =   14
      Top             =   90
      Width           =   5955
      Begin VB.TextBox txtAPPLICANT 
         Height          =   375
         Left            =   150
         TabIndex        =   4
         Top             =   420
         Width           =   5595
      End
      Begin VB.OptionButton optAPPLICANTID 
         Caption         =   "By Applicant ID"
         Height          =   285
         Left            =   2910
         TabIndex        =   6
         Top             =   1020
         Width           =   1935
      End
      Begin VB.OptionButton optLASTNAME 
         Caption         =   "Lastname"
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1020
         Width           =   1455
      End
      Begin MSComctlLib.ListView lsvAPPLICANT 
         Height          =   1845
         Left            =   150
         TabIndex        =   7
         Top             =   1830
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3254
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Name"
            Object.Width           =   8819
         EndProperty
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
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   3840
         Width           =   1860
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List Of Applicant"
         Height          =   240
         Index           =   8
         Left            =   150
         TabIndex        =   15
         Top             =   1500
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmAISEXAM_SCHED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SCHED_ID As Integer
Public DEL_SCHED_ID As Integer

Private Sub cboKINDofEXAM_Change()
    Dim rsTmp As ADODB.Recordset
    
    If Not lsvAPPEXAM.ListItems.Count = 0 Then
        '''''update the Type of Exam on Examinee
        lsvAPPEXAM.ListItems.Clear
        lsvAPPLICANT.ListItems.Clear
    End If
End Sub

Private Sub EnbledFrame(COND As Boolean)
    fmeAPP.Enabled = COND
    FmeSCHED.Enabled = COND
    fmeMAIN.Enabled = Not COND
End Sub

Private Sub cmdCancel_Click()
    Call EnbledFrame(False)
    lsvAPPEXAM.ListItems.Clear
    lsvSCHED.ListItems.Clear
    cboKINDofEXAM.SetFocus
    
    gconDMIS.Execute ("Delete From HRMS_APPLICANT_EXAM_SCHED Where Exam_Sched_ID = " & SCHED_ID & "")
    Call DisplayApplicantWhoGonnaExam
End Sub

Private Sub cmdEDIT_EXAM_Click()
    'frmMain.tbMENU.Enabled = False
    frmAISEXAM_SCHED.Enabled = False
    frmAISEXAM_EDIT.Show
End Sub

Private Sub cmdEXIT_Click()
    Unload Me
End Sub

Private Sub cmdSAVE_Click()
    Call EnbledFrame(False)
    lsvAPPEXAM.ListItems.Clear
    lsvSCHED.ListItems.Clear
    
    Call GenerateNewExamCTR
End Sub

Private Sub cmdSETTIME_Click()
    Dim VDTPDateOfExam As String, vcboTIMEofEXAM As String
    
    VDTPDateOfExam = N2Str2Null(DTPDATEofEXAM)
    vcboTIMEofEXAM = N2Str2Null(cboTIMEofEXAM)
    
    Call EnbledFrame(True)
    txtAPPLICANT.Text = ""
    txtAPPLICANT.SetFocus
    
    gconDMIS.Execute ("Delete From HRMS_EXAM_SCHEDULE Where EXAM_SCHED_ID = " & SCHED_ID & "")
    gconDMIS.Execute ("Delete From HRMS_APPLICANT_EXAM_SCHED Where EXam_sched_ID = " & SCHED_ID & "")
    
    gconDMIS.Execute ("Insert Into HRMS_EXAM_SCHEDULE Values(" & SCHED_ID & "," & Right(cboKINDofEXAM, 3) & _
        "," & VDTPDateOfExam & "," & vcboTIMEofEXAM & ")")
            
    Call DisplayApplicantScheduleThatDate
    Call EnbledFrame(True)
End Sub

Private Sub DisplayApplicantScheduleThatDate()
    Dim rsTmp As ADODB.Recordset, rsExam As ADODB.Recordset, rsPER As ADODB.Recordset
    Dim rsTYPE As ADODB.Recordset
    Dim VDATE  As String, VTIME As String
    Dim ITEM As ListItem
    
    VDATE = N2Str2Null(DTPDATEofEXAM)
    VTIME = N2Str2Null(cboTIMEofEXAM)
    
    Set rsTmp = GetRS("Select * From HRMS_EXAM_SCHEDULE Where DateOfExam = " & VDATE & _
        " And TimeOfExam = " & VTIME & "")
    
    lsvSCHED.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        DEL_SCHED_ID = rsTmp!EXAM_SCHED_ID
        'A
        Set rsExam = GetRS("Select Applicant_ID From HRMS_APPLICANT_EXAM_SCHED Where Exam_Sched_ID = " & _
            rsTmp!EXAM_SCHED_ID & "")

        If Not (rsExam.BOF And rsExam.EOF) Then
            Do While Not rsExam.EOF
                'B
                Set rsPER = GetRS("Select FirstName,LastName From HRMS_APPLICANT_PERSONAL Where Applicant_ID = " & _
                    rsExam!APPLICANT_ID & "")
                
                If Not (rsPER.BOF And rsPER.EOF) Then
                    Set ITEM = lsvSCHED.ListItems.Add(, , rsExam!APPLICANT_ID)
                    ITEM.SubItems(1) = rsPER!LastName & "," & rsPER!FirstName
                    
                    'C
                    Set rsTYPE = GetRS("Select ExamDescription From HRMS_ExamType Where ExamID = " & _
                        rsTmp!EXAMTYPE & "")
                    
                    If Not (rsTYPE.BOF And rsTYPE.EOF) Then
                        ITEM.SubItems(2) = rsTYPE!ExamDescription
                    End If
                    'C
                End If
                'B
                rsExam.MoveNext
            Loop
        End If
        'A
    End If
    
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    optLASTNAME.Value = True
        
    Call GenerateNewExamCTR
    Call FillTypeOfExam
    Call FillCBOTime(cboTIMEofEXAM)
End Sub

Private Sub GenerateNewExamCTR()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetRS("Select EXAM_SCHED_ID From HRMS_EXAM_SCHEDULE Order By EXAM_SCHED_ID ASC")
    
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            SCHED_ID = rsTmp!EXAM_SCHED_ID
            
            rsTmp.MoveNext
        Loop
    End If
    
    SCHED_ID = SCHED_ID + 1
End Sub

Function FillTypeOfExam()
    Dim rsTmp As ADODB.Recordset
    Dim SZERO As String
    
    Set rsTmp = GetRS("Select * From HRMS_ExamType Order By ExamDescription ASC")
    
    cboKINDofEXAM.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            If Len(rsTmp!EXAMID) = 1 Then SZERO = "00"
            If Len(rsTmp!EXAMID) = 2 Then SZERO = "0"
        
            cboKINDofEXAM.AddItem rsTmp!ExamDescription & " - " & SZERO & rsTmp!EXAMID
            
            rsTmp.MoveNext
        Loop
    End If
    
    cboKINDofEXAM.ListIndex = 0
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetRS("Select Exam_Sched_ID From HRMS_APPLICANT_EXAM_SCHED Where Exam_Sched_ID = " & SCHED_ID & "")
    If (rsTmp.BOF And rsTmp.EOF) Then
        gconDMIS.Execute ("Delete From HRMS_EXAM_SCHEDULE Where EXAM_SCHED_ID = " & SCHED_ID & "")
    End If
End Sub

Private Sub lsvAPPEXAM_DblClick()
    Dim INDEX As Integer
    
    If Not lsvAPPEXAM.ListItems.Count = 0 Then
        INDEX = lsvAPPEXAM.SelectedItem.INDEX
        
        With lsvAPPEXAM
            gconDMIS.Execute ("Delete From HRMS_APPLICANT_EXAM_SCHED Where ID = " & _
                .ListItems(INDEX).Text & " and EXAM_SCHED_ID = " & SCHED_ID & "")
            
            Call DisplayApplicantWhoGonnaExam
        End With
    End If
End Sub

Private Sub lsvAPPLICANT_DblClick()
    Dim INDEX As Integer
    Dim TAKEN_ALREADY As Boolean
    Dim CONFLICT_ON_TIME As Boolean
        
    If Not lsvAPPLICANT.ListItems.Count = 0 Then
        INDEX = lsvAPPLICANT.SelectedItem.INDEX
        
        With lsvAPPLICANT
            '-------------------------------------------------- APPLICANT ID----------, EXAM TYPE
            TAKEN_ALREADY = CheckifApplicantAlreadytakeThatExam(.ListItems(INDEX).Text, Right(cboKINDofEXAM, 3))

            If TAKEN_ALREADY = False Then
                '------------------------------------------APPLICANT ID
                CONFLICT_ON_TIME = CheckifNoConflictOnTime(.ListItems(INDEX).Text)
                
                If CONFLICT_ON_TIME = False Then
                    '--------------------------------------------APPLICANT ID
                    Call SaveToTheListofApplicantWillTaketheExam(.ListItems(INDEX).Text)
                Else
                    MsgBox "Applicant Exam Schedule Conflict", vbInformation, "Schedule of Exam"
                    lsvAPPLICANT.SetFocus
                End If
            Else
                MsgBox "Exam Type Already been schedule", vbInformation, "Schedule of Exam"
                txtAPPLICANT.SetFocus
            End If
        End With
    End If
End Sub

Private Sub SaveToTheListofApplicantWillTaketheExam(APP_ID As Integer)
    gconDMIS.Execute ("Insert Into HRMS_APPLICANT_EXAM_SCHED Values(" & APP_ID & "," & _
        SCHED_ID & "," & 0 & ",'" & "No Result Yet" & "')")
        
    Call DisplayApplicantWhoGonnaExam
End Sub

Private Sub DisplayApplicantWhoGonnaExam()
    Dim rsTmp As ADODB.Recordset, rsPER As ADODB.Recordset
    Dim ITEM As ListItem
    
    Set rsTmp = GetRS("Select * From HRMS_APPLICANT_EXAM_SCHED Where Exam_Sched_ID = " & SCHED_ID & "")
    
    lsvAPPEXAM.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvAPPEXAM.ListItems.Add(, , rsTmp!APPLICANT_ID)
            Set rsPER = GetRS("Select FirstName,LastName From HRMS_APPLICANT_PERSONAL Where  Applicant_ID = " & rsTmp!APPLICANT_ID & "")
            
            If Not (rsPER.BOF And rsPER.EOF) Then
                ITEM.SubItems(1) = rsPER!LastName & "," & rsPER!FirstName
            End If
            Set rsPER = Nothing
            
            rsTmp.MoveNext
        Loop
    End If
    
End Sub

'                                APPLICANT_ID,
Function CheckifNoConflictOnTime(APP_ID As Integer) As Boolean
    Dim rsTmp As ADODB.Recordset, rsSCHED As ADODB.Recordset
    
    Set rsTmp = GetRS("Select * From HRMS_APPLICANT_EXAM_SCHED Where Applicant_ID = " & APP_ID & "")
    
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set rsSCHED = GetRS("Select * From HRMS_EXAM_SCHEDULE Where EXAM_SCHED_ID = " & rsTmp!EXAM_SCHED_ID & "")
            
            If Not (rsSCHED.BOF And rsSCHED.EOF) Then
                If rsSCHED!DATEofEXAM = CDate(DTPDATEofEXAM) And rsSCHED!TimeOfExam = cboTIMEofEXAM Then
                    CheckifNoConflictOnTime = True
                    Exit Function
                Else
                    CheckifNoConflictOnTime = False
                End If
            End If
            
            rsTmp.MoveNext
        Loop
    Else
        CheckifNoConflictOnTime = False
    End If
    
End Function

'                                            APPLICANT ID , EXAM TYPE
Function CheckifApplicantAlreadytakeThatExam(ID As Integer, EXAMTYPE As Integer) As Boolean
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetRS("Select * From HRMS_APPLICANT_EXAM_SCHED Where Applicant_ID = " & ID & "")
    
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            '-------------------------------------------------------------EXAM SCHEDULE ID---, EXAM TYPE
            CheckifApplicantAlreadytakeThatExam = CheckIfExamAlreadyTaken(rsTmp!EXAM_SCHED_ID, EXAMTYPE)
            
            If CheckifApplicantAlreadytakeThatExam = True Then Exit Function
            rsTmp.MoveNext
        Loop
    Else
        CheckifApplicantAlreadytakeThatExam = False
    End If
End Function

Function CheckIfExamAlreadyTaken(EXAMSCHED As Integer, EXAMTYPE As Integer) As Boolean
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetRS("Select ExamType From HRMS_EXAM_SCHEDULE Where EXAM_SCHED_ID = " & EXAMSCHED & "")
    
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        If rsTmp!EXAMTYPE = EXAMTYPE Then
            CheckIfExamAlreadyTaken = True
        Else
            CheckIfExamAlreadyTaken = False
        End If
    Else
        CheckIfExamAlreadyTaken = False
    End If
End Function

Private Sub lsvAPPLICANT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lsvAPPLICANT_DblClick
End Sub

Private Sub lsvSCHED_DblClick()
    Dim INDEX As Long
    
    If Not lsvSCHED.ListItems.Count = 0 Then
        INDEX = CLng(lsvSCHED.SelectedItem.INDEX)
        With lsvSCHED
            If MsgBox("Remove Applicant on the Exam List", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
                gconDMIS.Execute ("Delete From HRMS_APPLICANT_EXAM_SCHED Where Applicant_ID = " & _
                    CInt(.ListItems(INDEX).Text) & " And EXAM_SCHED_ID = " & DEL_SCHED_ID & "")
                
                Call DisplayApplicantScheduleThatDate
            End If
        End With
    End If
End Sub

Private Sub lsvSCHED_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lsvSCHED_DblClick
End Sub

Private Sub optAPPLICANTID_Click()
    Call txtAPPLICANT_Change
End Sub

Private Sub optLASTNAME_Click()
    Call txtAPPLICANT_Change
End Sub

Private Sub txtAPPLICANT_Change()
    Dim Sql As String, Keyword As String
    Dim rsTmp As ADODB.Recordset
    Dim ITEM As ListItem
    
    Keyword = txtAPPLICANT.Text
    If optAPPLICANTID.Value Then
        Set rsTmp = GetRS("Select Applicant_ID,LastName,FirstName From HRMS_APPLICANT_PERSONAL Where Hired = '" & "NO" & _
            "' And Applicant_ID Like '" & Keyword & "%' Order By Applicant_ID ASC")
    Else
        Set rsTmp = GetRS("Select Applicant_ID,LastName,FirstName From HRMS_APPLICANT_PERSONAL Where Hired = '" & "NO" & _
            "' And LastName Like '" & Keyword & "%' Order By LastName ASC")
    End If
            
    lsvAPPLICANT.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvAPPLICANT.ListItems.Add(, , rsTmp!APPLICANT_ID)
            ITEM.SubItems(1) = Null2String(rsTmp!LastName) & "," & Null2String(rsTmp!FirstName)
        
            rsTmp.MoveNext
        Loop
    Else
        lsvAPPLICANT.ListItems.Clear
    End If
End Sub
