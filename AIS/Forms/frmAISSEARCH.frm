VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAISSEARCH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13905
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAISSEARCH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   13905
   Begin VB.Frame FmeSEARCH 
      Caption         =   "Search By"
      Height          =   3045
      Left            =   120
      TabIndex        =   20
      Top             =   3870
      Width           =   13695
      Begin VB.ComboBox cboPOSITION 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
      Begin VB.ComboBox cboSCHOOL 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   9840
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1710
         Width           =   3135
      End
      Begin VB.ComboBox cboCITY 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   9840
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   3105
      End
      Begin VB.ComboBox cboAGE 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1650
         Width           =   1365
      End
      Begin Crystal.CrystalReport rptSEARCH 
         Left            =   7650
         Top             =   2280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdEXIT 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   12750
         Picture         =   "frmAISSEARCH.frx":01CA
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exit Window"
         Top             =   2160
         Width           =   765
      End
      Begin VB.CommandButton cmdPRINT 
         Caption         =   "&Print"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   12000
         Picture         =   "frmAISSEARCH.frx":071C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Print Details"
         Top             =   2160
         Width           =   765
      End
      Begin VB.CheckBox chkBOX 
         Height          =   435
         Index           =   0
         Left            =   6510
         TabIndex        =   2
         Top             =   210
         Width           =   255
      End
      Begin VB.CommandButton cmdSEARCH 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   11250
         Picture         =   "frmAISSEARCH.frx":0CC0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Find Record"
         Top             =   2160
         Width           =   765
      End
      Begin VB.CheckBox chkBOX 
         Height          =   435
         Index           =   4
         Left            =   13110
         TabIndex        =   10
         Top             =   210
         Width           =   255
      End
      Begin VB.ComboBox cboFIELDS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   9840
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1230
         Width           =   3135
      End
      Begin VB.ComboBox cboGENDER 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   690
         Width           =   2055
      End
      Begin VB.ComboBox cboDEGREE 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   9840
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   3165
      End
      Begin VB.CheckBox chkBOX 
         Height          =   435
         Index           =   5
         Left            =   13110
         TabIndex        =   12
         Top             =   690
         Width           =   255
      End
      Begin VB.CheckBox chkBOX 
         Height          =   435
         Index           =   3
         Left            =   4170
         TabIndex        =   8
         Top             =   1620
         Width           =   255
      End
      Begin VB.CheckBox chkBOX 
         Height          =   435
         Index           =   2
         Left            =   5790
         TabIndex        =   6
         Top             =   1170
         Width           =   255
      End
      Begin VB.CheckBox chkBOX 
         Height          =   435
         Index           =   6
         Left            =   13110
         TabIndex        =   14
         Top             =   1170
         Width           =   255
      End
      Begin VB.CheckBox chkBOX 
         Height          =   435
         Index           =   7
         Left            =   13110
         TabIndex        =   16
         Top             =   1650
         Width           =   255
      End
      Begin VB.CheckBox chkBOX 
         Height          =   435
         Index           =   1
         Left            =   4860
         TabIndex        =   4
         Top             =   660
         Width           =   255
      End
      Begin VB.ComboBox cboCSTATUS 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1170
         Width           =   2985
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Position Desired"
         Height          =   240
         Index           =   0
         Left            =   1020
         TabIndex        =   28
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Gender"
         Height          =   240
         Index           =   11
         Left            =   1890
         TabIndex        =   27
         Top             =   810
         Width           =   690
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "School Name"
         Height          =   240
         Index           =   10
         Left            =   8460
         TabIndex        =   26
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Field Of Study"
         Height          =   240
         Index           =   9
         Left            =   8310
         TabIndex        =   25
         Top             =   1320
         Width           =   1410
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Age: Above"
         Height          =   240
         Index           =   8
         Left            =   1410
         TabIndex        =   24
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Highest Educ. Attainment"
         Height          =   240
         Index           =   7
         Left            =   7200
         TabIndex        =   23
         Top             =   870
         Width           =   2535
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Applicant Address"
         Height          =   240
         Index           =   1
         Left            =   7950
         TabIndex        =   22
         Top             =   330
         Width           =   1770
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Civil Status"
         Height          =   240
         Index           =   5
         Left            =   1440
         TabIndex        =   21
         Top             =   1260
         Width           =   1125
      End
   End
   Begin MSComctlLib.ListView lsvFILTER 
      Height          =   3765
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   6641
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Full Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Gender"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Civil Status"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Age"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Educ. Attaintment"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Study Fields"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Address"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "School name"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmAISSEARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report_Filter                                                     As String
Dim ConditionCount                                                    As Integer
Dim ConditionArray(8)                                                 As Boolean
Dim SQL                                                               As String

Function LOOP_CONDITION() As String
    '    Dim X As Integer, Y As Integer
    '    Dim SQL_RESULT As String
    '    Dim CONT_NEXT As Boolean
    LOOP_CONDITION = ""
    Dim CONDITION                                                     As String
    Dim X                                                             As Integer
    Dim FirstCondition                                                As Integer
    FirstCondition = 9

    For X = 0 To 7
        If chkBOX(X).Value = 1 Then
            ConditionArray(X) = True
        Else
            ConditionArray(X) = False
        End If
        If FirstCondition = 9 And ConditionArray(X) = True Then FirstCondition = X
    Next
    For X = 0 To 7
        If ConditionArray(X) = True And X = FirstCondition Then
            CONDITION = " WHERE " & Return_Condition(X)
        Else
            If ConditionArray(X) = True And X <> FirstCondition Then
                CONDITION = " AND " & Return_Condition(X)
            Else
                CONDITION = ""
            End If
        End If
        LOOP_CONDITION = LOOP_CONDITION & CONDITION
    Next

    '    For X = 0 To 7
    '        CONT_NEXT = False
    '        SQL_RESULT = ""
    '        For Y = 0 To X
    '            Call LOOPING1(X, Y, SQL_RESULT, CONT_NEXT)
    '
    '            If CONT_NEXT = True Then GoTo NEXT_ONE
    '        Next
    '
    'NEXT_ONE:
    '        Sql = Sql & SQL_RESULT
    '    Next
End Function

Function Return_Condition(XXX As Integer) As String
    If XXX = 0 Then Return_Condition = ("PositionDesired = '" & Trim(cboPosition) & "'")
    If XXX = 1 Then Return_Condition = ("GENDER = '" & Trim(cboGENDER) & "'")
    If XXX = 2 Then Return_Condition = ("CivilStatus = '" & Trim(cboCSTATUS) & "'")
    If XXX = 3 Then Return_Condition = ("Age >= " & cboAGE & "")
    If XXX = 4 Then Return_Condition = ("City = '" & Trim(cboCITY) & "'")
    If XXX = 5 Then Return_Condition = ("HighestLevel1 = '" & Trim(cboDEGREE) & "'")
    If XXX = 6 Then Return_Condition = ("StudyFields1 = '" & Trim(cboFIELDS) & "'")
    If XXX = 7 Then Return_Condition = ("SchoolName1 = '" & Trim(cboSCHOOL) & "'")
End Function

Function REPORT_LOOP_CONDITION() As String
    REPORT_LOOP_CONDITION = ""
    Dim CONDITION                                                     As String
    Dim X                                                             As Integer
    Dim FirstCondition                                                As Integer
    FirstCondition = 9
    For X = 0 To 7
        If chkBOX(X).Value = 1 Then                           'TRUE
            ConditionArray(X) = True
        Else                                                  'FALSE
            ConditionArray(X) = False
        End If
        If FirstCondition = 9 And ConditionArray(X) = True Then FirstCondition = X
    Next
    For X = 0 To 7
        If ConditionArray(X) = True And X = FirstCondition Then
            CONDITION = " " & Report_Condition(X)
        Else
            If ConditionArray(X) = True And X <> FirstCondition Then
                CONDITION = " AND " & Report_Condition(X)
            Else
                CONDITION = ""
            End If
        End If
        REPORT_LOOP_CONDITION = REPORT_LOOP_CONDITION & CONDITION
    Next

End Function

Function Report_Condition(XXX As Integer) As String
    If XXX = 0 Then Report_Condition = ("{HRMS_APPLICANT_PERSONAL.PositionDesired} = '" & Trim(cboPosition) & "'")
    If XXX = 1 Then Report_Condition = ("{HRMS_APPLICANT_PERSONAL.GENDER} = '" & Trim(cboGENDER) & "'")
    If XXX = 2 Then Report_Condition = ("{HRMS_APPLICANT_PERSONAL.CivilStatus} = '" & Trim(cboCSTATUS) & "'")
    If XXX = 3 Then Report_Condition = ("{HRMS_APPLICANT_PERSONAL.Age} >= " & cboAGE & "")
    If XXX = 4 Then Report_Condition = ("{HRMS_APPLICANT_PERSONAL.City} = '" & Trim(cboCITY) & "'")
    If XXX = 5 Then Report_Condition = ("{HRMS_APPLICANT_PERSONAL.HighestLevel1} = '" & Trim(cboDEGREE) & "'")
    If XXX = 6 Then Report_Condition = ("{HRMS_APPLICANT_PERSONAL.StudyFields1} = '" & Trim(cboFIELDS) & "'")
    If XXX = 7 Then Report_Condition = ("{HRMS_APPLICANT_PERSONAL.SchoolName1} = '" & Trim(cboSCHOOL) & "'")
End Function

Sub FillCBOAGE()
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select Distinct Age From HRMS_APPLICANT_PERSONAL Order by Age ASC")
    cboAGE.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboAGE.AddItem RSTMP!AGE
            RSTMP.MoveNext
        Loop
    End If
    cboAGE.ListIndex = 0
End Sub

Sub BackToDefault(cbo As ComboBox, COND As Boolean)
    cbo.Enabled = COND
    If COND Then cbo.BackColor = vbWhite
    If Not COND Then cbo.BackColor = &H8000000C
End Sub

Sub FIllCbo(cbo As ComboBox, STR As String)
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select DISTINCT " & STR & " As Field1 From HRMS_APPLICANT_PERSONAL Where " & STR & _
                    " IS NOT NULL Order By " & STR & " ASC")
    cbo.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cbo.AddItem Null2String(RSTMP!Field1)
            RSTMP.MoveNext
        Loop
    End If
    'Updating Code   >>   Ashish Piya
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub

Private Sub chkBOX_Click(Index As Integer)
    Select Case Index
        Case 0:
            If chkBOX(Index).Value = 1 Then
                cboPosition.Enabled = True
                cboPosition.BackColor = vbWhite
                On Error Resume Next
                cboPosition.SetFocus
            Else
                cboPosition.Enabled = False
                cboPosition.BackColor = &H8000000C
            End If

        Case 1:
            If chkBOX(Index).Value = 1 Then
                Call BackToDefault(cboGENDER, True)
                On Error Resume Next
                cboGENDER.SetFocus
            Else
                Call BackToDefault(cboGENDER, False)
                cboGENDER.BackColor = &H8000000C
            End If

        Case 2:
            If chkBOX(Index).Value = 1 Then
                Call BackToDefault(cboCSTATUS, True)
                On Error Resume Next
                cboCSTATUS.SetFocus
            Else
                Call BackToDefault(cboCSTATUS, False)
                cboCSTATUS.BackColor = &H8000000C
            End If

        Case 3:
            If chkBOX(Index).Value = 1 Then
                Call BackToDefault(cboAGE, True)
                On Error Resume Next
                cboAGE.SetFocus
            Else
                Call BackToDefault(cboAGE, False)
                cboAGE.BackColor = &H8000000C
            End If

        Case 4:
            If chkBOX(Index).Value = 1 Then
                Call BackToDefault(cboCITY, True)
                On Error Resume Next
                cboCITY.SetFocus
            Else
                Call BackToDefault(cboCITY, False)
            End If

        Case 5:
            If chkBOX(Index).Value = 1 Then
                Call BackToDefault(cboDEGREE, True)
                On Error Resume Next
                cboDEGREE.SetFocus
            Else
                Call BackToDefault(cboDEGREE, False)
            End If

        Case 6:
            If chkBOX(Index).Value = 1 Then
                Call BackToDefault(cboFIELDS, True)
                On Error Resume Next
                cboFIELDS.SetFocus
            Else
                Call BackToDefault(cboFIELDS, False)
            End If

        Case 7:
            If chkBOX(Index).Value = 1 Then
                Call BackToDefault(cboSCHOOL, True)
                On Error Resume Next
                cboSCHOOL.SetFocus
            Else
                Call BackToDefault(cboSCHOOL, False)
            End If
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim Position As String, GENDER As String, CIVILSTATUS As String, AGE As String, ADDRESS As String
    Dim DEGREE As String, FIELDS As String, SCHOOLNAME                As String
    Dim FILTER                                                        As String

    On Error GoTo Errorcode:
    frmMain.MousePointer = 11

    lsvFILTER.Enabled = False
    Call cmdSEARCH_Click
    If Not lsvFILTER.ListItems.count = 0 Then
        If chkBOX(0).Value = 1 Then Position = Trim(cboPosition)
        If chkBOX(1).Value = 1 Then GENDER = Trim(cboGENDER)
        If chkBOX(2).Value = 1 Then CIVILSTATUS = Trim(cboCSTATUS)
        If chkBOX(3).Value = 1 Then AGE = Trim(cboAGE)
        If chkBOX(4).Value = 1 Then ADDRESS = Trim(cboCITY)
        If chkBOX(5).Value = 1 Then DEGREE = Trim(cboDEGREE)
        If chkBOX(6).Value = 1 Then FIELDS = Trim(cboFIELDS)
        If chkBOX(7).Value = 1 Then SCHOOLNAME = Trim(cboSCHOOL)

        rptSEARCH.Formulas(0) = "Position = '" & Position & "'"
        rptSEARCH.Formulas(1) = "Gender = '" & GENDER & "'"
        rptSEARCH.Formulas(2) = "Status = '" & CIVILSTATUS & "'"
        rptSEARCH.Formulas(3) = "Age = '" & AGE & "'"
        rptSEARCH.Formulas(4) = "Address = '" & ADDRESS & "'"
        rptSEARCH.Formulas(5) = "Degree = '" & DEGREE & "'"
        rptSEARCH.Formulas(6) = "Fields = '" & FIELDS & "'"
        rptSEARCH.Formulas(7) = "SchoolName = '" & SCHOOLNAME & "'"
        rptSEARCH.Formulas(8) = "PrintedBy = '" & "K u T 0" & "'"

        On Error GoTo ERROR

        FILTER = REPORT_LOOP_CONDITION()
        Call PrintSQLReport(rptSEARCH, AIS_REPORT_PATH & "SearchReport.rpt", FILTER, AIS_REPORT_Connection, 1)
    End If
    lsvFILTER.Enabled = True
    frmMain.MousePointer = 0
    Exit Sub

ERROR:
    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub cmdSEARCH_Click()
    Dim RSTMP As ADODB.Recordset, rsALT                               As ADODB.Recordset
    Dim ITEM                                                          As ListItem
    Dim vtxtPOS As String, vcboGENDER As String, vcboSTATUS           As String
    Dim vtxtADD As String, vtxtDEGREE As String, vtxtFIELDS As String, vtxtSNAME As String
    Dim vtxtAGE                                                       As String
    Dim SQL_POS As String, SQL_GENDER As String, SQL_STATUS As String, SQL_AGE As String
    Dim SQL_ADD As String, SQL_EDU As String, SQL_FIELDS As String, SQL_SNAME As String

    On Error GoTo Errorcode:
    frmMain.MousePointer = 11

    SQL = "Select * From HRMS_APPLICANT_PERSONAL"

    'Call LOOPING(Sql)
    SQL = SQL & LOOP_CONDITION()
    SQL = SQL & " Order by Applicant_ID ASC"

    On Error GoTo ERROR
    Set RSTMP = gconDMIS.Execute(SQL)

    lsvFILTER.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        cmdPrint.Enabled = True

        Do While Not RSTMP.EOF
            Set ITEM = lsvFILTER.ListItems.Add(, , RSTMP!APPLICANT_ID)
            ITEM.SubItems(1) = Null2String(RSTMP!lastname & ", " & RSTMP!FIRSTNAME)
            ITEM.SubItems(2) = Null2String(RSTMP!GENDER)
            ITEM.SubItems(3) = Null2String(RSTMP!CIVILSTATUS)
            ITEM.SubItems(4) = Null2String(RSTMP!AGE)
            ITEM.SubItems(5) = Null2String(RSTMP!HighestLevel1)
            ITEM.SubItems(6) = Null2String(RSTMP!StudyFields1)
            ITEM.SubItems(7) = Null2String(RSTMP!ADDRESS)
            ITEM.SubItems(8) = Null2String(RSTMP!SchoolName1)

            RSTMP.MoveNext
        Loop
    Else
        MsgBox "No Applicant(s) Found", vbInformation, "Search Applicant"
        cmdPrint.Enabled = False
    End If

    frmMain.MousePointer = 0
    Exit Sub

ERROR:
    Beep
    lsvFILTER.ListItems.Clear
    Exit Sub
    ShowVBError
    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

'Sub LOOPING1(X As Integer, Y As Integer, SQL_RESULT As String, CONT_NEXT As Boolean)
'    If Not Y = X Then
'        If chkBOX(Y).Value = 1 Then
'            If Y = X Then                       'Where
'                If X = 0 Then Call LOPPING2(X, SQL_RESULT, " Where PositionDesired Like '%" & Trim(txtPOS) & "%'", CONT_NEXT)
'                If X = 1 Then Call LOPPING2(X, SQL_RESULT, " Where GENDER = '" & Trim(cboGENDER) & "'", CONT_NEXT)
'                If X = 2 Then Call LOPPING2(X, SQL_RESULT, " Where CivilStatus = '" & Trim(cboCSTATUS) & "'", CONT_NEXT)
'                If X = 3 Then Call LOPPING2(X, SQL_RESULT, " Where Age >= " & txtAGE & "", CONT_NEXT)
'                If X = 4 Then Call LOPPING2(X, SQL_RESULT, " Where Address Like '%" & Trim(txtADD) & "%'", CONT_NEXT)
'                If X = 5 Then Call LOPPING2(X, SQL_RESULT, " Where HighestLevel1 = '" & Trim(cboDEGREE) & "'", CONT_NEXT)
'                If X = 6 Then Call LOPPING2(X, SQL_RESULT, " Where StudyFields1 = '" & Trim(cboFIELDS) & "'", CONT_NEXT)
'                If X = 7 Then Call LOPPING2(X, SQL_RESULT, " Where SchoolName1 Like '%" & Trim(txtSchool) & "%'", CONT_NEXT)
'            Else                                'AND
'                If X = 1 Then Call LOPPING2(X, SQL_RESULT, " And GENDER = '" & Trim(cboGENDER) & "'", CONT_NEXT)
'                If X = 2 Then Call LOPPING2(X, SQL_RESULT, " And CivilStatus = '" & Trim(cboCSTATUS) & "'", CONT_NEXT)
'                If X = 3 Then Call LOPPING2(X, SQL_RESULT, " And Age >= " & txtAGE & "", CONT_NEXT)
'                If X = 4 Then Call LOPPING2(X, SQL_RESULT, " And Address Like '%" & Trim(txtADD) & "%'", CONT_NEXT)
'                If X = 5 Then Call LOPPING2(X, SQL_RESULT, " And HighestLevel1 = '" & Trim(cboDEGREE) & "'", CONT_NEXT)
'                If X = 6 Then Call LOPPING2(X, SQL_RESULT, " And StudyFields1 = '" & Trim(cboFIELDS) & "'", CONT_NEXT)
'                If X = 7 Then Call LOPPING2(X, SQL_RESULT, " And SchoolName1 Like '%" & Trim(txtSchool) & "%'", CONT_NEXT)
'            End If
'        End If
'    Else                                        'DEPENDS
'        If X = 0 Then Call LOPPING2(X, SQL_RESULT, " Where PositionDesired Like '%" & Trim(txtPOS) & "%'", CONT_NEXT)
'        If X = 1 Then Call LOPPING2(X, SQL_RESULT, " Where GENDER = '" & Trim(cboGENDER) & "'", CONT_NEXT)
'        If X = 2 Then Call LOPPING2(X, SQL_RESULT, " Where CivilStatus = '" & Trim(cboCSTATUS) & "'", CONT_NEXT)
'        If X = 3 Then Call LOPPING2(X, SQL_RESULT, " Where Age >= " & txtAGE & "", CONT_NEXT)
'        If X = 4 Then Call LOPPING2(X, SQL_RESULT, " Where Address Like '%" & Trim(txtADD) & "%'", CONT_NEXT)
'        If X = 5 Then Call LOPPING2(X, SQL_RESULT, " Where HighestLevel1 = '" & Trim(cboDEGREE) & "'", CONT_NEXT)
'        If X = 6 Then Call LOPPING2(X, SQL_RESULT, " Where StudyFields1 = '" & Trim(cboFIELDS) & "'", CONT_NEXT)
'        If X = 7 Then Call LOPPING2(X, SQL_RESULT, " Where SchoolName1 Like '%" & Trim(txtSchool) & "%'", CONT_NEXT)
'    End If
'End Sub
'=================================================================================================================
'Sub LOPPING2(X As Integer, SQL_RESULT As String, SQL_STMT As String, CONT_NEXT As Boolean)
'    If chkBOX(X).Value = 1 Then
'        SQL_RESULT = SQL_STMT
'        CONT_NEXT = True
'    End If
'    If chkBOX(X).Value = 0 Then SQL_RESULT = ""
'End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    frmMain.MousePointer = 11

    Call FIllCbo(cboGENDER, "Gender")
    Call FIllCbo(cboCSTATUS, "CivilStatus")
    Call FIllCbo(cboAGE, "Age")
    Call FIllCbo(cboPosition, "PositionDesired")
    Call FIllCbo(cboCITY, "City")
    Call FIllCbo(cboFIELDS, "StudyFields1")
    Call FIllCbo(cboDEGREE, "HighestLevel1")
    Call FIllCbo(cboSCHOOL, "SchoolName1")

    chkBOX(0).Value = 1
    Call cmdSEARCH_Click
    frmMain.MousePointer = 0
End Sub

Private Sub lsvFILTER_DblClick()
    Dim Index                                                         As Integer
    If Not lsvFILTER.ListItems.count = 0 Then
        Index = lsvFILTER.SelectedItem.Index
        frmAISApplications.Show
        frmAISApplications.tbcApplication.SelectedItem = 0
        APPLICANT_ID = lsvFILTER.ListItems(Index)
        Call frmAISApplications.DisplayAllInformation
        On Error Resume Next
        frmAISApplications.SetFocus
    End If
End Sub

