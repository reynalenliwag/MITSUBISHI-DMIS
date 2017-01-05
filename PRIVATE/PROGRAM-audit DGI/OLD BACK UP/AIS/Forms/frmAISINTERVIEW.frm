VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAISINTERVIEW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule Of Interview"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAISINTERVIEW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   9765
   Begin XtremeSuiteControls.TabControl tbcINTERVIEW 
      Height          =   9060
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9795
      _Version        =   655364
      _ExtentX        =   17277
      _ExtentY        =   15981
      _StockProps     =   64
      Appearance      =   9
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Schedule Interview"
      Item(0).Tooltip =   "Schedule Interview"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "fmeEXAM"
      Item(1).Caption =   "Reschedule Interview"
      Item(1).Tooltip =   "Reschedule Interview"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "Frame2"
      Begin VB.Frame fmeEXAM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8580
         Left            =   -69910
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   9555
         Begin MSComctlLib.ListView LsvTIME 
            Height          =   7590
            Left            =   90
            TabIndex        =   1
            ToolTipText     =   "Pls First set a date.."
            Top             =   720
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   13388
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
            MouseIcon       =   "frmAISINTERVIEW.frx":058A
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
            Height          =   7590
            Left            =   2250
            TabIndex        =   2
            Top             =   720
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   13388
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
            Left            =   2250
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
            Format          =   20643841
            CurrentDate     =   39143
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
            Left            =   90
            TabIndex        =   12
            Top             =   8355
            Width           =   3945
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Of Interview"
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
            Left            =   750
            TabIndex        =   11
            Top             =   360
            Width           =   1380
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8580
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   9555
         Begin MSComctlLib.ListView lsvTIME1 
            Height          =   7590
            Left            =   90
            TabIndex        =   4
            Top             =   720
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   13388
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
            Height          =   7590
            Left            =   2250
            TabIndex        =   5
            Top             =   720
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   13388
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
            Left            =   2250
            TabIndex        =   3
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
            Format          =   20643841
            CurrentDate     =   39143
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
            TabIndex        =   9
            Top             =   8355
            Width           =   2505
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Of Interview"
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
            Left            =   750
            TabIndex        =   8
            Top             =   360
            Width           =   1380
         End
      End
   End
End
Attribute VB_Name = "frmAISINTERVIEW"
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

Function FillSchedule1()
    Dim Item                                                          As ListItem
    Dim rsTmp                                                         As ADODB.Recordset
    Dim X                                                             As Integer
    Dim MYCOLOR                                                       As Long
    Dim TMP_ID                                                        As Long

    MYCOLOR = 16711680                      'Blue

    lsvACT1.ListItems.Clear
    For X = 1 To 32
        Set rsTmp = gconDMIS.Execute("Select * From HRMS_INTERVIEW_SCHEDULE Where FromTime < = " & _
                                   X & " And ToTime > = " & X & _
                                   " And DateOfInterview = '" & dtpDATE1 & "'")

        If Not (rsTmp.BOF And rsTmp.EOF) Then
            If Not TMP_ID = rsTmp!INT_ID Then
                MYCOLOR = ChangeColor(MYCOLOR)
            End If

            TMP_ID = rsTmp!INT_ID

            Set Item = lsvACT1.ListItems.Add(, , X)
            Item.SubItems(1) = Null2String(Trim(rsTmp!InterviewDescription))
            Item.SubItems(2) = rsTmp!INT_ID

            lsvACT1.ListItems(X).ListSubItems(1).ForeColor = MYCOLOR
        Else
            Set Item = lsvACT1.ListItems.Add(, , X)
            Item.SubItems(1) = ""
            Item.SubItems(2) = ""
        End If
    Next
End Function

Function ClickInterviewAndReSchedDate()
    Call dtpDATE_Change
    Call dtpDATE1_Change
End Function

Function FillSchedule()
    Dim Item                                                          As ListItem
    Dim rsTmp                                                         As New ADODB.Recordset
    Dim X                                                             As Integer
    Dim MYCOLOR                                                       As Long
    Dim TMP_ID                                                        As Long

    MYCOLOR = 16711680                      'Blue

    lsvACT.ListItems.Clear
    If CDate(dtpDATE) < Date Then
        MsgBox "Cannot Schedule An Interview that is Already Passed", vbInformation, "Schedule Of Interview"
        dtpDATE.SetFocus
        Exit Function
    End If

    For X = 1 To 32
        Set rsTmp = gconDMIS.Execute("Select * From HRMS_INTERVIEW_SCHEDULE Where FromTime < = " & _
                                   X & " And ToTime > = " & X & _
                                   " And DateOfInterview = '" & dtpDATE & "'")

        If Not (rsTmp.BOF And rsTmp.EOF) Then
            If Not TMP_ID = rsTmp!INT_ID Then
                MYCOLOR = ChangeColor(MYCOLOR)
            End If

            TMP_ID = rsTmp!INT_ID

            Set Item = lsvACT.ListItems.Add(, , X)
            Item.SubItems(1) = Null2String(rsTmp!InterviewDescription)
            Item.SubItems(2) = rsTmp!INT_ID

            lsvACT.ListItems(X).ListSubItems(1).ForeColor = MYCOLOR
        Else
            Set Item = lsvACT.ListItems.Add(, , X)
            Item.SubItems(1) = ""
            Item.SubItems(2) = ""
        End If
    Next
End Function

Sub FillListViewTime()
    Dim Item                                                          As ListItem
    Dim ITEM1                                                         As ListItem
    Dim rsTmp                                                         As ADODB.Recordset

    LsvTIME.ListItems.Clear
    lsvTIME1.ListItems.Clear
    Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME4 Order By Time_ID ASC")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LsvTIME.ListItems.Add(, , rsTmp!Set_Time)
            Item.SubItems(1) = rsTmp!Time_ID

            Set ITEM1 = lsvTIME1.ListItems.Add(, , rsTmp!Set_Time)
            ITEM1.SubItems(1) = rsTmp!Time_ID

            rsTmp.MoveNext
        Loop
    End If
End Sub

Sub FIllFromTime_RESCHED()
    Dim rsTmp                                                         As ADODB.Recordset
    Dim SZERO                                                         As String

    Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME3 Order By TIME_ID ASC")
    frmAISINTERVIEW_RESCHED.cboFTIME.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            If Not rsTmp!Time_ID = 17 And Not rsTmp!Time_ID = 34 Then
                If Len(rsTmp!Time_ID) = 1 Then SZERO = "00"
                If Len(rsTmp!Time_ID) = 2 Then SZERO = "0"

                frmAISINTERVIEW_RESCHED.cboFTIME.AddItem rsTmp!Set_Time & Space(10) & SZERO & rsTmp!Time_ID
            End If
            rsTmp.MoveNext
        Loop
    End If
    frmAISINTERVIEW_RESCHED.cboFTIME.ListIndex = 0
End Sub

Sub DisplayScheduleTime(ID As Integer)
    Dim rsTmp                                                         As ADODB.Recordset
    Dim TIME1 As String, TIME2                                        As String
    Dim SZERO                                                         As String

    Set rsTmp = gconDMIS.Execute("Select FromTime,ToTime From HRMS_INTERVIEW_SCHEDULE Where INT_ID = " & ID & "")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        If rsTmp!FROMTIME < 17 Then
            TIME1 = GetTime_TMP(rsTmp!FROMTIME)

            frmAISINTERVIEW_RESCHED.lblINFO(2).Caption = TIME1
            If Len(rsTmp!FROMTIME) = 1 Then SZERO = "00"
            If Len(rsTmp!FROMTIME) = 2 Then SZERO = "0"
            frmAISINTERVIEW_RESCHED.cboFTIME.Text = TIME1 & Space(10) & SZERO & rsTmp!FROMTIME

            TIME2 = GetTime_TMP(rsTmp!ToTime + 1)
            frmAISINTERVIEW_RESCHED.lblINFO(3).Caption = TIME2
            If Len(rsTmp!ToTime + 1) = 1 Then SZERO = "00"
            If Len(rsTmp!ToTime + 1) = 2 Then SZERO = "0"

            frmAISINTERVIEW_RESCHED.cboTTIME.Text = TIME2 & Space(10) & SZERO & rsTmp!ToTime + 1
        End If

        If rsTmp!FROMTIME > 16 Then
            TIME1 = GetTime_TMP(rsTmp!FROMTIME + 1)

            frmAISINTERVIEW_RESCHED.lblINFO(2).Caption = TIME1
            If Len(rsTmp!FROMTIME + 1) = 1 Then SZERO = "00"
            If Len(rsTmp!FROMTIME + 1) = 2 Then SZERO = "0"
            frmAISINTERVIEW_RESCHED.cboFTIME.Text = TIME1 & Space(10) & SZERO & rsTmp!FROMTIME + 1

            TIME2 = GetTime_TMP(rsTmp!ToTime + 2)
            frmAISINTERVIEW_RESCHED.lblINFO(3).Caption = TIME2
            If Len(rsTmp!ToTime + 2) = 1 Then SZERO = "00"
            If Len(rsTmp!ToTime + 2) = 2 Then SZERO = "0"

            frmAISINTERVIEW_RESCHED.cboTTIME.Text = TIME2 & Space(10) & SZERO & rsTmp!ToTime + 2
        End If
    End If
End Sub

Private Sub dtpDATE_Change()
    Call FillSchedule
End Sub

Private Sub dtpDATE1_Change()
    Call FillSchedule1
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    Call FillListViewTime
    'dtpDATE.Day = 1
    'dtpDATE.Year = Year(Date)
    'dtpDATE.Month = Month(Date)
    'dtpDATE.Day = Day(Date)
    dtpDATE.Value = Date

    'dtpDATE1.Day = 1
    'dtpDATE1.Year = Year(Date)
    'dtpDATE1.Month = Month(Date)
    'dtpDATE1.Day = Day(Date)
    dtpDATE1.Value = Date
    Call ClickInterviewAndReSchedDate

    tbcINTERVIEW.SelectedItem = 0
End Sub

Private Sub lsvACT_Click()
    CLICK_LSV = "ACT"
End Sub

Private Sub lsvACT_DblClick()
    Dim TIME1 As String, TIME2                                        As String
    Dim rsTmp                                                         As ADODB.Recordset

    If Not lsvACT.ListItems.count = 0 Then
        If Not CLICK_LSV = "TIME" Then INDEX = lsvACT.SelectedItem.INDEX

        With lsvACT
            If Not CInt(.ListItems(INDEX).Text) > 32 Then
                If .ListItems(INDEX).SubItems(1) = "" Then    'SCHEDULE NEW
                    frmAISINTERVIEW.Enabled = False
                    frmAISINTERVIEW_SET.Show
                    frmAISINTERVIEW_SET.lblFTIME.Caption = LsvTIME.ListItems(INDEX).SubItems(1)

                    Call FillSetInterviewTime(CInt(.ListItems(INDEX).Text), INDEX)
                    On Error Resume Next
                    frmAISINTERVIEW_SET.txtDesc.SetFocus
                Else                                          'VIEW OR EDIT OF GRADE
                    frmAISINTERVIEW.Enabled = False
                    frmAISINTERVIEW_VIEW.Show
                    frmAISINTERVIEW_VIEW.lblSCHED_ID.Caption = CInt(.ListItems(INDEX).SubItems(2))
                    frmAISINTERVIEW_VIEW.lblINFO(0).Caption = dtpDATE
                    frmAISINTERVIEW_VIEW.lblINFO(1).Caption = lsvACT.ListItems(INDEX).SubItems(1)

                    Set rsTmp = gconDMIS.Execute("Select INT_ID,FromTime,ToTime From HRMS_INTERVIEW_SCHEDULE Where INT_ID = " & .ListItems(INDEX).SubItems(2) & "")
                    If Not (rsTmp.BOF And rsTmp.EOF) Then
                        If rsTmp!FROMTIME < 17 Then
                            TIME1 = GetTime_TMP(rsTmp!FROMTIME)
                            frmAISINTERVIEW_VIEW.lblINFO(2).Caption = TIME1
                            TIME2 = GetTime_TMP(rsTmp!ToTime + 1)
                            frmAISINTERVIEW_VIEW.lblINFO(3).Caption = TIME2
                        End If
                        If rsTmp!FROMTIME > 16 Then
                            TIME1 = GetTime_TMP(rsTmp!FROMTIME + 1)
                            frmAISINTERVIEW_VIEW.lblINFO(2).Caption = TIME1
                            TIME2 = GetTime_TMP(rsTmp!ToTime + 2)
                            frmAISINTERVIEW_VIEW.lblINFO(3).Caption = TIME2
                        End If

                        frmAISINTERVIEW_VIEW.lblEXAMID.Caption = rsTmp!INT_ID
                    End If

                    Call frmAISINTERVIEW_VIEW.DisplayList(CLng(.ListItems(INDEX).SubItems(2)))
                End If
            Else
            End If
        End With
    End If
End Sub

Private Sub FillSetInterviewTime(FROMTIME As Integer, INDEX As Long)
    Dim rsTmp As ADODB.Recordset, rsTIME                              As ADODB.Recordset
    Dim INDEX2                                                        As Integer
    Dim SZERO                                                         As String

    If INDEX <= 16 Then Set rsTIME = gconDMIS.Execute("Select Set_Time From HRMS_TIME3 Where Time_ID = " & INDEX & "")
    If INDEX >= 17 Then Set rsTIME = gconDMIS.Execute("Select Set_Time From HRMS_TIME3 Where Time_ID = " & INDEX + 1 & "")
    If Not (rsTIME.BOF And rsTIME.EOF) Then frmAISINTERVIEW_SET.lblFROMTIME.Caption = rsTIME!Set_Time

    frmAISINTERVIEW_SET.cboTIMEofEXAM.Clear
    If INDEX < 17 Then
        INDEX2 = INDEX
        Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME3 Where Time_ID < = " & 17 & _
                                   " And Time_ID > " & INDEX & " Order By Time_ID ASC")
        If Not (rsTmp.BOF And rsTmp.EOF) Then
            Do While Not rsTmp.EOF
                If lsvACT.ListItems(INDEX2).SubItems(1) = "" Then
                    SZERO = ""
                    If Len(LsvTIME.ListItems(INDEX2 + 1).SubItems(1)) = 1 Then SZERO = "0"

                    frmAISINTERVIEW_SET.cboTIMEofEXAM.AddItem rsTmp!Set_Time & "          " & SZERO & LsvTIME.ListItems(INDEX2).SubItems(1)
                Else
                    GoTo CONT1
                End If

                INDEX2 = INDEX2 + 1
                rsTmp.MoveNext
            Loop
        End If
    End If

    If INDEX > 16 Then
        INDEX2 = INDEX
        Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME3 Where Time_ID > " & 17 & _
                                   " And Time_ID > " & INDEX + 1 & " Order By Time_ID ASC")
        If Not (rsTmp.BOF And rsTmp.EOF) Then
            Do While Not rsTmp.EOF
                If lsvACT.ListItems(INDEX2).SubItems(1) = "" Then
                    SZERO = ""
                    If Len(LsvTIME.ListItems(INDEX2).SubItems(1)) = 1 Then SZERO = "0"

                    frmAISINTERVIEW_SET.cboTIMEofEXAM.AddItem rsTmp!Set_Time & "          " & SZERO & LsvTIME.ListItems(INDEX2).SubItems(1)
                Else
                    GoTo CONT1
                End If

                INDEX2 = INDEX2 + 1
                rsTmp.MoveNext
            Loop
        End If
    End If
CONT1:

    frmAISINTERVIEW_SET.cboTIMEofEXAM.ListIndex = 0
End Sub

Private Sub lsvACT1_Click()
    CLICK_LSV1 = "ACT"
End Sub

Private Sub lsvACT1_DblClick()
    Dim rsTIME                                                        As ADODB.Recordset

    If Not lsvACT1.ListItems.count = 0 Then
        If CDate(Date) <= dtpDATE1 Then
            If Not CLICK_LSV1 = "TIME" Then INDEX1 = lsvACT1.SelectedItem.INDEX

            With lsvACT1
                If Not lsvACT1.ListItems(INDEX1).SubItems(1) = "" Then
                    frmAISINTERVIEW.Enabled = False
                    frmAISINTERVIEW_RESCHED.Show
                    frmAISINTERVIEW_RESCHED.lblSCHED_ID.Caption = .ListItems(INDEX1).SubItems(2)
                    frmAISINTERVIEW_RESCHED.oPENsCHEDULE

                    If Not CInt(.ListItems(INDEX1).Text) > 32 Then
                        frmAISINTERVIEW_RESCHED.lblDATE.Caption = dtpDATE1
                        frmAISINTERVIEW_RESCHED.lblFID.Caption = .ListItems(INDEX1).Text
                        frmAISINTERVIEW_RESCHED.lblSCHED_ID.Caption = .ListItems(INDEX1).SubItems(2)
                        frmAISINTERVIEW_RESCHED.lblINFO(0).Caption = .ListItems(INDEX1).SubItems(1)
                        frmAISINTERVIEW_RESCHED.txtDesc.Text = .ListItems(INDEX1).SubItems(1)
                        frmAISINTERVIEW_RESCHED.lblINFO(1).Caption = dtpDATE1

                        frmAISINTERVIEW_RESCHED.dtpDATE.Day = Day(dtpDATE1)
                        frmAISINTERVIEW_RESCHED.dtpDATE.Month = Month(dtpDATE1)
                        frmAISINTERVIEW_RESCHED.dtpDATE.YEAR = YEAR(dtpDATE1)

                        If INDEX1 < 17 Then Set rsTIME = gconDMIS.Execute("Select Set_Time From HRMS_TIME3 Where Time_ID = " & INDEX1 & "")
                        If INDEX1 > 16 Then Set rsTIME = gconDMIS.Execute("Select Set_Time From HRMS_TIME3 Where Time_ID = " & INDEX1 + 1 & "")

                        If Not (rsTIME.BOF And rsTIME.EOF) Then frmAISINTERVIEW_RESCHED.lblINFO(2).Caption = rsTIME!Set_Time

                        Call FIllFromTime_RESCHED

                        Call DisplayScheduleTime(.ListItems(INDEX1).SubItems(2))

                        On Error Resume Next
                        frmAISINTERVIEW_RESCHED.txtDesc.SetFocus
                    End If
                End If
            End With
        Else
            MsgBox "Interview Date Already Passed", vbInformation, "Reschedule Of Interview"
            If lsvACT1.ListItems.count > 0 And lsvACT1.Enabled = True Then: lsvACT1.SetFocus

        End If
    End If
End Sub

Private Sub LsvTIME_DblClick()
    If Not LsvTIME.ListItems.count = 0 Then
        CLICK_LSV = "TIME"
        INDEX = LsvTIME.SelectedItem.INDEX
        With LsvTIME
            Call lsvACT_DblClick
        End With
    End If
End Sub

Private Sub lsvTIME1_DblClick()
    If Not lsvTIME1.ListItems.count = 0 Then
        CLICK_LSV1 = "TIME"
        INDEX1 = lsvTIME1.SelectedItem.INDEX
        With lsvTIME1
            Call lsvACT1_DblClick
        End With
    End If
End Sub

