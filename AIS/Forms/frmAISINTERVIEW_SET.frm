VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAISINTERVIEW_SET 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Interview"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAISINTERVIEW_SET.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   7965
   Begin VB.Frame FmeSCHED 
      Caption         =   "Set Time"
      Height          =   2265
      Left            =   60
      TabIndex        =   15
      Top             =   0
      Width           =   7785
      Begin VB.TextBox txtDESC 
         Appearance      =   0  'Flat
         Height          =   900
         Left            =   1980
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   240
         Width           =   5625
      End
      Begin VB.CommandButton cmdCAN 
         Caption         =   "Cancel Set"
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
         Left            =   6630
         Picture         =   "frmAISINTERVIEW_SET.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancel/Set Interview"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.ComboBox cboTIMEofEXAM 
         Height          =   360
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1710
         Width           =   1635
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
         Left            =   5610
         Picture         =   "frmAISINTERVIEW_SET.frx":0C16
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Set Interview"
         Top             =   1320
         Width           =   1035
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
         Left            =   1170
         TabIndex        =   20
         Top             =   1860
         Width           =   690
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
         Left            =   990
         TabIndex        =   19
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label lblFROMTIME 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1980
         TabIndex        =   18
         Top             =   1230
         Width           =   2145
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   2
         Left            =   90
         TabIndex        =   17
         Top             =   330
         Width           =   1830
      End
      Begin VB.Label lblFTIME 
         BackColor       =   &H000000FF&
         Caption         =   " "
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.Frame FmeLIST 
      Caption         =   "Choose Applicants"
      Enabled         =   0   'False
      Height          =   5565
      Left            =   90
      TabIndex        =   9
      Top             =   2340
      Width           =   7785
      Begin VB.TextBox txtSEARCH 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1050
         TabIndex        =   4
         Top             =   330
         Width           =   5175
      End
      Begin MSComctlLib.ListView LsvAPP 
         Height          =   2745
         Left            =   150
         TabIndex        =   5
         Top             =   1530
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   4842
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
         Height          =   2775
         Left            =   4020
         TabIndex        =   8
         Top             =   1500
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   4895
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
         Height          =   795
         Left            =   6630
         Picture         =   "frmAISINTERVIEW_SET.frx":13A2
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exit Window"
         Top             =   4650
         Width           =   975
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
         Height          =   795
         Left            =   5670
         Picture         =   "frmAISINTERVIEW_SET.frx":18F4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save Entry"
         Top             =   4650
         Width           =   975
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
         Left            =   1050
         TabIndex        =   24
         Top             =   750
         Width           =   1200
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
         TabIndex        =   14
         Top             =   420
         Width           =   675
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
         TabIndex        =   13
         Top             =   1200
         Width           =   1890
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
         Left            =   4050
         TabIndex        =   12
         Top             =   1230
         Width           =   1800
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
         TabIndex        =   11
         Top             =   4350
         Width           =   1860
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
         Left            =   4050
         TabIndex        =   10
         Top             =   4320
         Width           =   2235
      End
   End
   Begin VB.Label lblPrevDesc 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   8400
      TabIndex        =   23
      Top             =   1110
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblPrevTime 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   8400
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblSCHED_ID 
      BackColor       =   &H000000FF&
      Caption         =   " "
      Height          =   735
      Left            =   8400
      TabIndex        =   21
      Top             =   180
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmAISINTERVIEW_SET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RECALL                                                            As String
Dim TITLE                                                             As String

Private Sub CheckForConflict()
    Dim rsTmp As ADODB.Recordset, rsSCHED                             As ADODB.Recordset

    Set rsTmp = gconDMIS.Execute("Select * From HRMS_APPLICANT_INTERVIEW_SCHEDULE ORder By Applicant_ID")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set rsSCHED = gconDMIS.Execute("Select * From HRMS_INTERVIEW_SCHEDULE Where INT_ID = " & CLng(rsTmp!INT_ID) & "")
            If Not (rsSCHED.BOF And rsSCHED.EOF) Then
                If CInt(lblFTIME.Caption) < rsSCHED!FROMTIME Then
                    If CInt(Right(cboTIMEofEXAM, 2)) >= rsSCHED!FROMTIME Then
                        gconDMIS.Execute ("Delete From HRMS_APPLICANT_INTERVIEW_SCHEDULE_TMP Where Applicant_ID = " & rsTmp!APPLICANT_ID & "")
                    End If
                End If
                If CInt(lblFTIME.Caption) = rsSCHED!FROMTIME Then
                    gconDMIS.Execute ("Delete From HRMS_APPLICANT_INTERVIEW_SCHEDULE_TMP Where Applicant_ID = " & rsTmp!APPLICANT_ID & "")
                End If
                If CInt(lblFTIME.Caption) > rsSCHED!FROMTIME Then
                    If CInt(lblFTIME.Caption) <= rsSCHED!ToTime Then
                        gconDMIS.Execute ("Delete From HRMS_APPLICANT_INTERVIEW_SCHEDULE_TMP Where Applicant_ID = " & rsTmp!APPLICANT_ID & "")
                    End If
                End If
            End If

            rsTmp.MoveNext
        Loop
    End If
End Sub

Private Sub DisplayChange()
    Dim rsTmp                                                         As ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Set rsTmp = gconDMIS.Execute("Select * From HRMS_APPLICANT_INTERVIEW_SCHEDULE_TMP")
    lsvList.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvList.ListItems.Add(, , rsTmp!APPLICANT_ID)
            ITEM.SubItems(1) = rsTmp!FULLNAME

            rsTmp.MoveNext
        Loop
    End If
End Sub

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:17
Private Sub cmdCAN_Click()
    cmdSET.Caption = "CHANGE"
    cmdCAN.Visible = False
    txtDesc.Text = lblPrevDesc.Caption
    cboTIMEofEXAM.Text = lblPrevTime.Caption

    txtDesc.Enabled = False
    cboTIMEofEXAM.Enabled = False
    FmeLIST.Enabled = True
    txtSearch.Text = ""
    txtSearch.SetFocus
    On Error GoTo Errorcode:

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
    Dim DATEOFINTERVIEW                                               As String
    Dim INTERVIEW_DESC                                                As String
    Dim REMARKS As String, NOTE                                       As String
    Dim rsTmp                                                         As ADODB.Recordset

    On Error GoTo Errorcode:
    frmMain.MousePointer = 11

    If Not lsvList.ListItems.count = 0 Then
        If MsgBox("Save This Interview Schedule", vbQuestion + vbYesNo + vbDefaultButton1, "Are You Sure") = vbYes Then
            If txtDesc.Text = "" Then
                MsgBox "Enter a Interview Description", vbExclamation, "Schedule Of Interview"
                On Error Resume Next
                txtDesc.SetFocus
                Exit Sub
            End If

            INTERVIEW_DESC = N2Str2Null(Trim(txtDesc.Text))
            DATEOFINTERVIEW = N2Str2Null(frmAISINTERVIEW.dtpDate)
            REMARKS = N2Str2Null("")
            NOTE = N2Str2Null("")

            gconDMIS.Execute ("Insert Into HRMS_INTERVIEW_SCHEDULE Values(" & CInt(lblSCHED_ID) & _
                              "," & INTERVIEW_DESC & _
                              "," & DATEOFINTERVIEW & _
                              "," & CInt(lblFTIME) & _
                              "," & CInt(Right(cboTIMEofEXAM, 2)) & ")")

            Set rsTmp = gconDMIS.Execute("Select * From HRMS_APPLICANT_INTERVIEW_SCHEDULE_TMP Order By Applicant_ID ASC")
            If Not (rsTmp.BOF And rsTmp.EOF) Then
                Do While Not rsTmp.EOF
                    gconDMIS.Execute ("Insert Into HRMS_APPLICANT_INTERVIEW_SCHEDULE Values(" & rsTmp!APPLICANT_ID & _
                                      "," & CLng(lblSCHED_ID) & _
                                      "," & REMARKS & _
                                      "," & NOTE & ")")

                    rsTmp.MoveNext
                Loop
            End If

            Unload Me

            Call frmAISINTERVIEW.FillSchedule
            Call frmAISINTERVIEW.FillSchedule1
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

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:17
Private Sub cmdSET_Click()
    On Error GoTo Errorcode:

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

                lblPrevDesc.Caption = txtDesc.Text
                lblPrevTime.Caption = cboTIMEofEXAM.Text

                On Error Resume Next
                txtSearch.SetFocus
            End If
        Else
            MsgBox "Activity Descripion Cannot be Blank", vbInformation, "Schedule of Examination"
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
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3:
            If FmeLIST.Enabled Then
                txtSearch.Text = ""
                On Error Resume Next
                txtSearch.SetFocus
            End If

    End Select
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    gconDMIS.Execute ("Delete From HRMS_APPLICANT_INTERVIEW_SCHEDULE_TMP")
    Call txtsearch_Change
    Call GenerateNewSCHED_ID
End Sub

Private Sub GenerateNewSCHED_ID()
    Dim rsTmp                                                         As ADODB.Recordset

    lblSCHED_ID.Caption = 0
    Set rsTmp = gconDMIS.Execute("Select TOP 1 INT_ID FROM HRMS_INTERVIEW_SCHEDULE Order By INT_ID DESC")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        lblSCHED_ID.Caption = rsTmp!INT_ID
    End If
    lblSCHED_ID.Caption = lblSCHED_ID.Caption + 1

    gconDMIS.Execute ("Delete From HRMS_APPLICANT_EXAM_SCHEDULE_TMP")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAISINTERVIEW.Enabled = True
    On Error Resume Next
    frmAISINTERVIEW.SetFocus
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

Private Sub LsvAPP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call LsvAPP_DblClick
End Sub

Private Sub lsvLIST_DblClick()
    Dim INDEX                                                         As Integer
    If Not lsvList.ListItems.count = 0 Then
        INDEX = lsvList.SelectedItem.INDEX
        With lsvList
            If MsgBox("Remove This Applicant From the Interview List", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
                gconDMIS.Execute ("Delete From HRMS_APPLICANT_INTERVIEW_SCHEDULE_TMP Where Applicant_ID = " & CLng(.ListItems(INDEX).Text) & "")
                Call DisplayChange
            End If
        End With
    End If
End Sub

Private Sub lsvLIST_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lsvLIST_DblClick
End Sub

Private Sub txtsearch_Change()
    Dim Keyword                                                       As String
    Dim rsTmp                                                         As ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Keyword = Trim(txtSearch)
    Set rsTmp = gconDMIS.Execute("Select Applicant_ID,FirstName,Lastname From HRMS_APPLICANT_PERSONAL Where FirstName Like '%" & _
                                 Keyword & "%' Or Lastname Like '%" & Keyword & "%' Order by Applicant_ID ASC")
    LsvAPP.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = LsvAPP.ListItems.Add(, , rsTmp!APPLICANT_ID)
            ITEM.SubItems(1) = Null2String(rsTmp!lastname & ", " & rsTmp!FIRSTNAME)

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

    Set rsTmp = gconDMIS.Execute("Select INT_ID,REMARKS From HRMS_APPLICANT_INTERVIEW_SCHEDULE Where APPLICANT_ID = " & APP_ID & "")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set rsSCHED = gconDMIS.Execute("Select * From HRMS_INTERVIEW_SCHEDULE Where INT_ID = " & rsTmp!INT_ID & "")
            If Not (rsSCHED.BOF And rsSCHED.EOF) Then
                X_ID = rsSCHED!INT_ID
                t1 = rsSCHED!FROMTIME
                T2 = rsSCHED!ToTime

                If Null2String(rsTmp!REMARKS) = "" Then                     'REMARKS -> Already Interview
                    MsgBox "Applicant Already Scheduled to Interview", vbInformation, "Schedule of Exam"
                    If LsvAPP.ListItems.count > 0 And LsvAPP.Enabled = True Then: LsvAPP.SetFocus
                    GoTo ALREADY_TAKE
                ElseIf Null2String(rsTmp!REMARKS) = "Passed" Then           'REMARKS -> PASSED
                    MsgBox "Applicant Already Passed The Interview", vbInformation, "Schedule Of Exam"
                    If LsvAPP.ListItems.count > 0 And LsvAPP.Enabled = True Then: LsvAPP.SetFocus
                    GoTo ALREADY_TAKE
                Else                                                        'REMARKS -> FAILED or EXAM NOT YET TAKEN
                    GoTo JUMP1
                End If
JUMP1:
                '-----------------------------------------------------------------------------------------------------
                If rsSCHED!DATEOFINTERVIEW = frmAISINTERVIEW.dtpDate Then
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
            rsTmp.MoveNext
        Loop
        GoTo JUMP_ELSE
    Else
JUMP_ELSE:
        Set rsExam = gconDMIS.Execute("Select * From HRMS_APPLICANT_INTERVIEW_SCHEDULE_TMP Where APPLICANT_ID = " & APP_ID & "")
        If (rsExam.BOF And rsExam.EOF) Then
            gconDMIS.Execute ("Insert Into HRMS_APPLICANT_INTERVIEW_SCHEDULE_TMP Values(" & APP_ID & _
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
    NOTES = NOTES & " (" & TIME1 & " - " & TIME2 & ")"
    MsgBox NOTES, vbExclamation, "Schedule Of Exam"

    If LsvAPP.ListItems.count > 0 And LsvAPP.Enabled = True Then: LsvAPP.SetFocus
End Sub

