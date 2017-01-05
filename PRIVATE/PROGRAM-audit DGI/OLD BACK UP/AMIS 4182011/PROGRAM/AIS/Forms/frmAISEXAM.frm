VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAISEXAM 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exam"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAISEXAM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   9720
   Begin XtremeSuiteControls.TabControl tbcEXAM 
      Height          =   9060
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9735
      _Version        =   655364
      _ExtentX        =   17171
      _ExtentY        =   15981
      _StockProps     =   64
      Appearance      =   3
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   2
      Item(0).Caption =   "Schedule Exam"
      Item(0).Tooltip =   "Schedule Exam"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "fmeEXAM"
      Item(1).Caption =   "Reschedule Exam"
      Item(1).Tooltip =   "Reschedule Exam"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "Frame2"
      Begin VB.Frame Frame2 
         Height          =   8640
         Left            =   -69910
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   9555
         Begin VB.ComboBox cboEXAMTYPE1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5970
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   270
            Width           =   2800
         End
         Begin MSComctlLib.ListView lsvTIME1 
            Height          =   7620
            Left            =   90
            TabIndex        =   6
            Top             =   720
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   13441
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
               Text            =   "Time"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   9
            EndProperty
         End
         Begin MSComctlLib.ListView lsvACT1 
            Height          =   7620
            Left            =   2250
            TabIndex        =   7
            Top             =   720
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   13441
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
               Object.Width           =   9
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Activity Description"
               Object.Width           =   12347
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   9
            EndProperty
         End
         Begin MSComCtl2.DTPicker dtpDATE1 
            Height          =   345
            Left            =   1710
            TabIndex        =   4
            Top             =   270
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20250625
            CurrentDate     =   39154
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Exam"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kind Of Exam"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   4560
            TabIndex        =   15
            Top             =   390
            Width           =   1110
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Double Click to Reschedule"
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
            Index           =   3
            Left            =   90
            TabIndex        =   12
            Top             =   8385
            Width           =   2505
         End
      End
      Begin VB.Frame fmeEXAM 
         Height          =   8640
         Left            =   90
         TabIndex        =   9
         Top             =   360
         Width           =   9555
         Begin VB.ComboBox cboExamType 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5970
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   270
            Width           =   2800
         End
         Begin MSComctlLib.ListView LsvTIME 
            Height          =   7620
            Left            =   90
            TabIndex        =   2
            Top             =   720
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   13441
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
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmAISEXAM.frx":0ECA
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Time"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   9
            EndProperty
         End
         Begin MSComctlLib.ListView lsvACT 
            Height          =   7620
            Left            =   2250
            TabIndex        =   3
            Top             =   720
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   13441
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
               Object.Width           =   9
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Activity Desciption "
               Object.Width           =   12347
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   9
            EndProperty
         End
         Begin MSComCtl2.DTPicker dtpDATE 
            Height          =   345
            Left            =   1710
            TabIndex        =   0
            Top             =   270
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20250625
            CurrentDate     =   39154
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Exam"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   36
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kind Of Exam"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   20
            Left            =   4560
            TabIndex        =   13
            Top             =   390
            Width           =   1110
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Double Click to Add or View Exam Schedule"
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
            Left            =   120
            TabIndex        =   11
            Top             =   8385
            Width           =   3945
         End
      End
   End
End
Attribute VB_Name = "frmAISEXAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim INDEX As Long, INDEX1                                             As Long
Attribute INDEX1.VB_VarUserMemId = 1073938432
Dim CLICK_LSV As String, CLICK_LSV1                                   As String
Attribute CLICK_LSV.VB_VarUserMemId = 1073938434
Attribute CLICK_LSV1.VB_VarUserMemId = 1073938434

Sub FillExamType()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim SZERO                                                         As String

    Set RSTMP = gconDMIS.Execute("Select ExamID,ExamDescription From HRMS_ExamType Order By ExamID ASC")

    cboExamType.Clear
    cboEXAMTYPE1.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            If Len(RSTMP!EXAMID) = 1 Then SZERO = "00"
            If Len(RSTMP!EXAMID) = 2 Then SZERO = "0"

            cboExamType.AddItem RSTMP!ExamDescription & " - " & SZERO & RSTMP!EXAMID
            cboEXAMTYPE1.AddItem RSTMP!ExamDescription & " - " & SZERO & RSTMP!EXAMID
            RSTMP.MoveNext
        Loop
        cboExamType.ListIndex = 0
        cboEXAMTYPE1.ListIndex = 0
    End If

End Sub

Sub FillSchedule()
    Dim Item                                                          As ListItem
    Dim RSTMP                                                         As ADODB.Recordset
    Dim X                                                             As Integer
    Dim DATEofEXAM                                                    As String
    Dim MYCOLOR                                                       As Long
    Dim TMP_ID                                                        As Long

    DATEofEXAM = dtpDATE
    MYCOLOR = 16711680

    lsvACT.ListItems.Clear
    For X = 1 To 32
        Set RSTMP = gconDMIS.Execute("Select * From HRMS_EXAM_SCHEDULE Where FromTime < = " & _
                                   X & " And ToTime > = " & X & " And ExamID = " & CInt(Right(cboExamType, 3)) & _
                                   " And DateOfExam = '" & CDate(dtpDATE) & "'")

        If Not (RSTMP.BOF And RSTMP.EOF) Then
            'If Not X = 1 Then
            'lsvACT.ListItems(X).ListSubItems(1).ForeColor = MYCOLOR
            If Not TMP_ID = RSTMP!SCHED_ID Then
                MYCOLOR = ChangeColor(MYCOLOR)
            End If

            TMP_ID = RSTMP!SCHED_ID
            'End If

            Set Item = lsvACT.ListItems.Add(, , X)
            Item.SubItems(1) = RSTMP!ActivityDescription
            Item.SubItems(2) = RSTMP!SCHED_ID

            lsvACT.ListItems(X).ListSubItems(1).ForeColor = MYCOLOR
        Else
            Set Item = lsvACT.ListItems.Add(, , X)
            Item.SubItems(1) = ""
            Item.SubItems(2) = ""
        End If
    Next
End Sub

Sub FillSchedule1()
    Dim Item                                                          As ListItem
    Dim RSTMP                                                         As ADODB.Recordset
    Dim X                                                             As Integer
    Dim DATEofEXAM                                                    As String
    Dim MYCOLOR                                                       As Long
    Dim TMP_ID                                                        As Long

    DATEofEXAM = dtpDATE
    MYCOLOR = 16711680

    'For X = 1 To 16
    lsvACT1.ListItems.Clear
    For X = 1 To 32
        Set RSTMP = gconDMIS.Execute("Select * From HRMS_EXAM_SCHEDULE Where FromTime < = " & _
                                   X & " And ToTime > = " & X & " And ExamID = " & CInt(Right(cboEXAMTYPE1, 3)) & _
                                   " And DateOfExam = '" & CDate(dtpDATE1) & "'")

        If Not (RSTMP.BOF And RSTMP.EOF) Then
            'If Not X = 1 Then
            'lsvACT.ListItems(X).ListSubItems(1).ForeColor = MYCOLOR
            If Not TMP_ID = RSTMP!SCHED_ID Then
                MYCOLOR = ChangeColor(MYCOLOR)
            End If

            TMP_ID = RSTMP!SCHED_ID
            'End If

            Set Item = lsvACT1.ListItems.Add(, , X)
            Item.SubItems(1) = RSTMP!ActivityDescription
            Item.SubItems(2) = RSTMP!SCHED_ID

            lsvACT1.ListItems(X).ListSubItems(1).ForeColor = MYCOLOR
            'Set ITEM = lsvACT1.ListItems.Add(, , X)
            'ITEM.SubItems(1) = rsTmp!ActivityDescription
            'ITEM.SubItems(2) = rsTmp!SCHED_ID
        Else
            Set Item = lsvACT1.ListItems.Add(, , X)
            Item.SubItems(1) = ""
            Item.SubItems(2) = ""
        End If
    Next
End Sub

Sub FillListViewTime()
    Dim Item                                                          As ListItem
    Dim ITEM1                                                         As ListItem
    Dim RSTMP                                                         As ADODB.Recordset

    LsvTIME.ListItems.Clear
    lsvTIME1.ListItems.Clear
    'Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME2 Order By Time_ID ASC")
    Set RSTMP = gconDMIS.Execute("Select * From HRMS_TIME4 Order By Time_ID ASC")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set Item = LsvTIME.ListItems.Add(, , RSTMP!Set_Time)
            Item.SubItems(1) = RSTMP!Time_ID

            Set ITEM1 = lsvTIME1.ListItems.Add(, , RSTMP!Set_Time)
            ITEM1.SubItems(1) = RSTMP!Time_ID

            RSTMP.MoveNext
        Loop
    End If
End Sub

Sub DisplayScheduleTime(ID As Integer)
    Dim RSTMP                                                         As ADODB.Recordset
    Dim TIME1 As String, TIME2                                        As String
    Dim SZERO                                                         As String

    Set RSTMP = gconDMIS.Execute("Select FromTime,ToTime From HRMS_EXAM_SCHEDULE Where Sched_ID = " & ID & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        'If rsTmp!FROMTIME < 9 Then
        If RSTMP!FROMTIME < 17 Then
            TIME1 = GetTime_TMP(RSTMP!FROMTIME)

            frmAISEXAM_RESCHED.lblINFO(2).Caption = TIME1
            If Len(RSTMP!FROMTIME) = 1 Then SZERO = "00"
            If Len(RSTMP!FROMTIME) = 2 Then SZERO = "0"
            'Debug.Print TIME1 & Space(10) & SZERO & rsTmp!FROMTIME
            frmAISEXAM_RESCHED.cboFTIME.Text = TIME1 & Space(10) & SZERO & RSTMP!FROMTIME

            TIME2 = GetTime_TMP(RSTMP!ToTime + 1)
            frmAISEXAM_RESCHED.lblINFO(3).Caption = TIME2
            If Len(RSTMP!ToTime + 1) = 1 Then SZERO = "00"
            If Len(RSTMP!ToTime + 1) = 2 Then SZERO = "0"
Debug.Print TIME2 & Space(10) & SZERO & RSTMP!ToTime + 1
            frmAISEXAM_RESCHED.cboTTIME.Text = TIME2 & Space(10) & SZERO & RSTMP!ToTime + 1
        End If

        'If rsTmp!FROMTIME > 8 Then
        If RSTMP!FROMTIME > 16 Then
            TIME1 = GetTime_TMP(RSTMP!FROMTIME + 1)

            frmAISEXAM_RESCHED.lblINFO(2).Caption = TIME1
            If Len(RSTMP!FROMTIME + 1) = 1 Then SZERO = "00"
            If Len(RSTMP!FROMTIME + 1) = 2 Then SZERO = "0"
            'Debug.Print TIME1 & Space(10) & SZERO & rsTmp!FROMTIME + 1
            frmAISEXAM_RESCHED.cboFTIME.Text = TIME1 & Space(10) & SZERO & RSTMP!FROMTIME + 1

            TIME2 = GetTime_TMP(RSTMP!ToTime + 2)
            frmAISEXAM_RESCHED.lblINFO(3).Caption = TIME2
            If Len(RSTMP!ToTime + 2) = 1 Then SZERO = "00"
            If Len(RSTMP!ToTime + 2) = 2 Then SZERO = "0"
            'Debug.Print TIME2 & Space(10) & SZERO & rsTmp!ToTIME + 2
            frmAISEXAM_RESCHED.cboTTIME.Text = TIME2 & Space(10) & SZERO & RSTMP!ToTime + 2
        End If
    End If
End Sub

Sub FIllFromTime_RESCHED()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim SZERO                                                         As String

    'Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME1 Order By TIME_ID ASC")

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_TIME3 Order By TIME_ID ASC")
    frmAISEXAM_RESCHED.cboFTIME.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            'If Not rsTmp!Time_ID = 9 And Not rsTmp!Time_ID = 18 Then

            If Not RSTMP!Time_ID = 17 And Not RSTMP!Time_ID = 34 Then
                If Len(RSTMP!Time_ID) = 1 Then SZERO = "00"
                If Len(RSTMP!Time_ID) = 2 Then SZERO = "0"

                'Debug.Print rsTmp!SETTIME & Space(10) & SZERO & rsTmp!Time_ID
                'frmAISEXAM_RESCHED.cboFTIME.AddItem rsTmp!SETTIME & Space(10) & SZERO & rsTmp!Time_ID

                frmAISEXAM_RESCHED.cboFTIME.AddItem RSTMP!Set_Time & Space(10) & SZERO & RSTMP!Time_ID
            End If
            RSTMP.MoveNext
        Loop
    End If
    frmAISEXAM_RESCHED.cboFTIME.ListIndex = 0
End Sub

Private Sub cboExamType_Change()
    Call FillSchedule
End Sub

Private Sub cboExamType_Click()
    Call FillSchedule
End Sub

Private Sub cboEXAMTYPE1_Change()
    Call FillSchedule1
End Sub

Private Sub cboEXAMTYPE1_Click()
    Call cboEXAMTYPE1_Change
End Sub

Private Sub dtpDATE_Change()
    Call cboExamType_Change
End Sub

Private Sub dtpDATE1_Change()
    Call cboEXAMTYPE1_Change
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    Call FillListViewTime
    Call FillExamType

    dtpDATE.Day = 1
    dtpDATE.Year = Year(Date)
    dtpDATE.Month = Month(Date)
    dtpDATE.Day = Day(Date)

    dtpDATE1.Day = 1
    dtpDATE1.Year = Year(Date)
    dtpDATE1.Month = Month(Date)
    dtpDATE1.Day = Day(Date)

    tbcEXAM.SelectedItem = 0
End Sub

Private Sub lsvACT_Click()
    CLICK_LSV = "ACT"
End Sub

Private Sub lsvACT_DblClick()
    Dim TIME1 As String, TIME2                                        As String
    Dim RSTMP                                                         As ADODB.Recordset

    If Not lsvACT.ListItems.Count = 0 Then
        If Not CLICK_LSV = "TIME" Then INDEX = lsvACT.SelectedItem.INDEX

        With lsvACT
            If Not CInt(.ListItems(INDEX).Text) > 32 Then
                If .ListItems(INDEX).SubItems(1) = "" Then    'SCHEDULE NEW
                    frmAISEXAM.Enabled = False
                    frmAISEXAM_SET.Show
                    frmAISEXAM_SET.lblFTIME.Caption = LsvTIME.ListItems(INDEX).SubItems(1)
                    frmAISEXAM_SET.lblExamType.Caption = Mid(cboExamType, 1, Len(cboExamType) - 6)

                    Call FillSetExamTime(CInt(.ListItems(INDEX).Text), INDEX)
                    On Error Resume Next
                    frmAISEXAM_SET.txtDesc.SetFocus
                Else                                          'VIEW OR EDIT
                    frmAISEXAM.Enabled = False
                    frmAISEXAM_VIEW.Show
                    frmAISEXAM_VIEW.lblSCHED_ID.Caption = CInt(.ListItems(INDEX).SubItems(2))
                    frmAISEXAM_VIEW.lblINFO(0).Caption = dtpDATE
                    frmAISEXAM_VIEW.lblINFO(1).Caption = Mid(cboExamType, 1, Len(cboExamType) - 6)
                    frmAISEXAM_VIEW.lblINFO(2).Caption = lsvACT.ListItems(INDEX).SubItems(1)

                    Set RSTMP = gconDMIS.Execute("Select ExamID,FromTime,ToTime From HRMS_EXAM_SCHEDULE Where SCHED_ID = " & .ListItems(INDEX).SubItems(2) & "")
                    If Not (RSTMP.BOF And RSTMP.EOF) Then
                        If RSTMP!FROMTIME < 17 Then
                            TIME1 = GetTime_TMP(RSTMP!FROMTIME)
                            frmAISEXAM_VIEW.lblINFO(3).Caption = TIME1
                            TIME2 = GetTime_TMP(RSTMP!ToTime + 1)
                            frmAISEXAM_VIEW.lblINFO(4).Caption = TIME2
                        End If
                        If RSTMP!FROMTIME > 16 Then
                            TIME1 = GetTime_TMP(RSTMP!FROMTIME + 1)
                            frmAISEXAM_VIEW.lblINFO(3).Caption = TIME1
                            TIME2 = GetTime_TMP(RSTMP!ToTime + 2)
                            frmAISEXAM_VIEW.lblINFO(4).Caption = TIME2
                        End If

                        frmAISEXAM_VIEW.lblEXAMID.Caption = RSTMP!EXAMID
                    End If

                    Call frmAISEXAM_VIEW.DisplayList(CLng(.ListItems(INDEX).SubItems(2)))
                End If
            Else
            End If
        End With
    End If
End Sub

Private Sub FillSetExamTime(FROMTIME As Integer, INDEX As Long)
    Dim RSTMP As ADODB.Recordset, rsTIME                              As ADODB.Recordset
    Dim INDEX2                                                        As Integer
    Dim SZERO                                                         As String

    '----------------------------------------------------------------------------------------------------------
    'If INDEX <= 8 Then Set rsTIME = gconDMIS.Execute("Select SetTime From HRMS_TIME1 Where Time_ID = " & INDEX & "")
    'If INDEX >= 9 Then Set rsTIME = gconDMIS.Execute("Select SetTime From HRMS_TIME1 Where Time_ID = " & INDEX + 1 & "")
    If INDEX <= 16 Then Set rsTIME = gconDMIS.Execute("Select Set_Time From HRMS_TIME3 Where Time_ID = " & INDEX & "")
    If INDEX >= 17 Then Set rsTIME = gconDMIS.Execute("Select Set_Time From HRMS_TIME3 Where Time_ID = " & INDEX + 1 & "")

    'If Not (rsTIME.BOF And rsTIME.EOF) Then frmAISEXAM_SET.lblFROMTIME.Caption = rsTIME!SETTIME
    If Not (rsTIME.BOF And rsTIME.EOF) Then frmAISEXAM_SET.lblFROMTIME.Caption = rsTIME!Set_Time
    '----------------------------------------------------------------------------------------------------------

    frmAISEXAM_SET.cboTIMEofEXAM.Clear
    If INDEX < 17 Then
        INDEX2 = INDEX
        'Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME1 Where Time_ID < = " & 9 & _
         '    " And Time_ID > " & INDEX & " Order By Time_ID ASC")
        Set RSTMP = gconDMIS.Execute("Select * From HRMS_TIME3 Where Time_ID < = " & 17 & _
                                   " And Time_ID > " & INDEX & " Order By Time_ID ASC")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            Do While Not RSTMP.EOF
                If lsvACT.ListItems(INDEX2).SubItems(1) = "" Then
                    SZERO = ""
                    If Len(LsvTIME.ListItems(INDEX2 + 1).SubItems(1)) = 1 Then SZERO = "0"

                    frmAISEXAM_SET.cboTIMEofEXAM.AddItem RSTMP!Set_Time & "          " & SZERO & LsvTIME.ListItems(INDEX2).SubItems(1)
                Else
                    GoTo CONT1
                End If

                INDEX2 = INDEX2 + 1
                RSTMP.MoveNext
            Loop
        End If
    End If

    If INDEX > 16 Then
        INDEX2 = INDEX
        'Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME1 Where Time_ID > " & 9 & _
         '    " And Time_ID > " & INDEX + 1 & " Order By Time_ID ASC")

        Set RSTMP = gconDMIS.Execute("Select * From HRMS_TIME3 Where Time_ID > " & 17 & _
                                   " And Time_ID > " & INDEX + 1 & " Order By Time_ID ASC")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            Do While Not RSTMP.EOF
                If lsvACT.ListItems(INDEX2).SubItems(1) = "" Then
                    SZERO = ""
                    If Len(LsvTIME.ListItems(INDEX2).SubItems(1)) = 1 Then SZERO = "0"

                    frmAISEXAM_SET.cboTIMEofEXAM.AddItem RSTMP!Set_Time & "          " & SZERO & LsvTIME.ListItems(INDEX2).SubItems(1)
                Else
                    GoTo CONT1
                End If

                INDEX2 = INDEX2 + 1
                RSTMP.MoveNext
            Loop
        End If
    End If
CONT1:

    frmAISEXAM_SET.cboTIMEofEXAM.ListIndex = 0
End Sub

Private Sub lsvACT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lsvACT_DblClick
End Sub

Private Sub lsvACT1_Click()
    CLICK_LSV1 = "ACT"
End Sub

Private Sub lsvACT1_DblClick()
    Dim rsTIME                                                        As ADODB.Recordset

    If Not lsvACT1.ListItems.Count = 0 Then
        If Date <= dtpDATE1 Then
            If Not CLICK_LSV1 = "TIME" Then INDEX1 = lsvACT1.SelectedItem.INDEX

            With lsvACT1
                If Not lsvACT1.ListItems(INDEX1).SubItems(1) = "" Then
                    frmAISEXAM.Enabled = False
                    frmAISEXAM_RESCHED.Show
                    frmAISEXAM_RESCHED.lblSCHED_ID.Caption = .ListItems(INDEX1).SubItems(2)
                    frmAISEXAM_RESCHED.oPENsCHEDULE

                    If Not CInt(.ListItems(INDEX1).Text) > 32 Then
                        frmAISEXAM_RESCHED.lblDATE.Caption = dtpDATE1
                        frmAISEXAM_RESCHED.lblFID.Caption = .ListItems(INDEX1).Text
                        frmAISEXAM_RESCHED.lblSCHED_ID.Caption = .ListItems(INDEX1).SubItems(2)
                        frmAISEXAM_RESCHED.lblINFO(0).Caption = .ListItems(INDEX1).SubItems(1)
                        frmAISEXAM_RESCHED.txtDesc.Text = .ListItems(INDEX1).SubItems(1)
                        frmAISEXAM_RESCHED.lblINFO(1).Caption = dtpDATE1
                        frmAISEXAM_RESCHED.dtpDATE.Day = Day(dtpDATE1)
                        frmAISEXAM_RESCHED.dtpDATE.Month = Month(dtpDATE1)
                        frmAISEXAM_RESCHED.dtpDATE.Year = Year(dtpDATE1)

                        'If INDEX1 <= 8 Then Set rsTIME = gconDMIS.Execute("Select SetTime From HRMS_TIME1 Where Time_ID = " & INDEX1 & "")
                        'If INDEX1 >= 9 Then Set rsTIME = gconDMIS.Execute("Select SetTime From HRMS_TIME1 Where Time_ID = " & INDEX1 + 1 & "")

                        If INDEX1 < 17 Then Set rsTIME = gconDMIS.Execute("Select Set_Time From HRMS_TIME3 Where Time_ID = " & INDEX1 & "")
                        If INDEX1 > 16 Then Set rsTIME = gconDMIS.Execute("Select Set_Time From HRMS_TIME3 Where Time_ID = " & INDEX1 + 1 & "")

                        If Not (rsTIME.BOF And rsTIME.EOF) Then frmAISEXAM_RESCHED.lblINFO(2).Caption = rsTIME!Set_Time

                        Call FIllFromTime_RESCHED

                        Call DisplayScheduleTime(.ListItems(INDEX1).SubItems(2))
                        On Error Resume Next
                        frmAISEXAM_RESCHED.txtDesc.SetFocus
                    End If
                End If
            End With
        Else
            MsgBox "Interview Date Already Passed", vbExclamation, "Schedule Of Interview"
            On Error Resume Next
            lsvACT1.SetFocus
        End If
    End If
End Sub

Private Sub lsvACT1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lsvACT1_DblClick
End Sub

Private Sub LsvTIME_DblClick()
    If Not LsvTIME.ListItems.Count = 0 Then
        CLICK_LSV = "TIME"
        INDEX = LsvTIME.SelectedItem.INDEX
        With LsvTIME
            Call lsvACT_DblClick
        End With
    End If
End Sub

Private Sub LsvTIME_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call LsvTIME_DblClick
End Sub

Private Sub lsvTIME1_DblClick()
    If Not lsvTIME1.ListItems.Count = 0 Then
        CLICK_LSV1 = "TIME"
        INDEX1 = lsvTIME1.SelectedItem.INDEX
        With lsvTIME1
            Call lsvACT1_DblClick
        End With
    End If
End Sub

Private Sub lsvTIME1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lsvTIME1_DblClick
End Sub

