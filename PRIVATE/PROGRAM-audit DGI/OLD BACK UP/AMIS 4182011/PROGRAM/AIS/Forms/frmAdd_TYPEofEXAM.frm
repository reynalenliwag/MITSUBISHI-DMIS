VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAISAdd_TYPEofEXAM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Type Of Exam"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
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
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAdd_TYPEofEXAM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7605
   Begin VB.PictureBox picDEPT 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   150
      ScaleHeight     =   1095
      ScaleWidth      =   7395
      TabIndex        =   20
      Top             =   4830
      Width           =   7425
      Begin VB.CommandButton cmdDEPT_EDIT 
         Caption         =   "EDIT"
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
         Left            =   6285
         TabIndex        =   2
         Top             =   660
         Width           =   975
      End
      Begin VB.CommandButton cmdDEPT_NEW 
         Caption         =   "NEW"
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
         Left            =   5325
         TabIndex        =   1
         Top             =   660
         Width           =   885
      End
      Begin VB.ComboBox cboDEPT 
         Height          =   360
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
         Width           =   5715
      End
      Begin VB.Label lblCAP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   210
         Width           =   1155
      End
   End
   Begin VB.PictureBox picEDIT_DEPT 
      Height          =   1575
      Left            =   6180
      ScaleHeight     =   1515
      ScaleWidth      =   5655
      TabIndex        =   17
      Top             =   4710
      Visible         =   0   'False
      Width           =   5715
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "SAVE"
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
         Left            =   3480
         Picture         =   "frmAdd_TYPEofEXAM.frx":09AA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   660
         Width           =   975
      End
      Begin VB.CommandButton cndCANCEL 
         Caption         =   "CANCEL"
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
         Left            =   4500
         Picture         =   "frmAdd_TYPEofEXAM.frx":104A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox txtDEPT 
         Height          =   360
         Left            =   210
         TabIndex        =   14
         Top             =   210
         Width           =   5235
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   6720
      ScaleHeight     =   495
      ScaleWidth      =   7395
      TabIndex        =   25
      Top             =   4110
      Width           =   7425
      Begin VB.TextBox txtID 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   360
         Left            =   1500
         TabIndex        =   26
         Top             =   60
         Width           =   1275
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   240
         Index           =   0
         Left            =   1050
         TabIndex        =   27
         Top             =   180
         Width           =   210
      End
   End
   Begin VB.PictureBox picEXAM 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   8730
      ScaleHeight     =   2685
      ScaleWidth      =   2415
      TabIndex        =   24
      Top             =   1860
      Width           =   2445
      Begin MSComctlLib.ListView lsvDEPT_EXAM 
         Height          =   2235
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3942
         View            =   3
         Arrange         =   1
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
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Exam Type"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   88
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
         Index           =   3
         Left            =   60
         TabIndex        =   28
         Top             =   2370
         Width           =   2235
      End
   End
   Begin VB.PictureBox picEDIT_EXAM 
      Height          =   2295
      Left            =   210
      ScaleHeight     =   2235
      ScaleWidth      =   6075
      TabIndex        =   29
      Top             =   150
      Visible         =   0   'False
      Width           =   6135
      Begin VB.PictureBox picSAVE 
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
         Height          =   885
         Left            =   4440
         ScaleHeight     =   885
         ScaleWidth      =   1485
         TabIndex        =   34
         Top             =   1200
         Visible         =   0   'False
         Width           =   1485
         Begin VB.CommandButton cmdEXAM_CANCEL 
            Caption         =   "&Cancel"
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
            Left            =   720
            Picture         =   "frmAdd_TYPEofEXAM.frx":15C6
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Cancel"
            Top             =   60
            Width           =   735
         End
         Begin VB.CommandButton cmdEXAM_SAVE 
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
            Left            =   0
            Picture         =   "frmAdd_TYPEofEXAM.frx":1B42
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Save Exam Type"
            Top             =   60
            Width           =   735
         End
      End
      Begin VB.TextBox txtMAX 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1545
         TabIndex        =   9
         Top             =   1260
         Width           =   1665
      End
      Begin VB.TextBox txtPGRADE 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1545
         TabIndex        =   8
         Top             =   810
         Width           =   1665
      End
      Begin VB.TextBox txtMIN 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1545
         TabIndex        =   10
         Top             =   1680
         Width           =   1665
      End
      Begin VB.TextBox txtEXAMDESC 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1530
         TabIndex        =   7
         Top             =   420
         Width           =   4395
      End
      Begin VB.CommandButton cmdADD_to_DEPT 
         Caption         =   "ADD EXAM TO DEPARTMENT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         TabIndex        =   11
         Top             =   3180
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label lblCAP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         Height          =   240
         Index           =   20
         Left            =   1065
         TabIndex        =   33
         Top             =   1320
         Width           =   390
      End
      Begin VB.Label lblCAP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passing Grade"
         Height          =   240
         Index           =   23
         Left            =   90
         TabIndex        =   32
         Top             =   900
         Width           =   1380
      End
      Begin VB.Label lblCAP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
         Height          =   240
         Index           =   2
         Left            =   1125
         TabIndex        =   31
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label lblCAP 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Type"
         Height          =   240
         Index           =   28
         Left            =   390
         TabIndex        =   30
         Top             =   480
         Width           =   1080
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   6555
         _Version        =   655364
         _ExtentX        =   11562
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "       "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         ForeColor       =   0
      End
   End
   Begin VB.PictureBox picENTRY 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   150
      ScaleHeight     =   3555
      ScaleWidth      =   7335
      TabIndex        =   18
      Top             =   90
      Width           =   7365
      Begin Crystal.CrystalReport rptEXAM 
         Left            =   480
         Top             =   3060
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComctlLib.ListView lsvEXAM 
         Height          =   2325
         Left            =   120
         TabIndex        =   3
         Top             =   150
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   4101
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
            Text            =   "Exam Type"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Passing"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Min"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Max"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.PictureBox picADD 
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
         Left            =   4800
         ScaleHeight     =   825
         ScaleWidth      =   2445
         TabIndex        =   19
         Top             =   2640
         Width           =   2445
         Begin VB.CommandButton cmdEXAM_EXIT 
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
            Left            =   1650
            Picture         =   "frmAdd_TYPEofEXAM.frx":21E2
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exit Window"
            Top             =   0
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
            Left            =   870
            Picture         =   "frmAdd_TYPEofEXAM.frx":2734
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Print Type Of Exam"
            Top             =   0
            Width           =   795
         End
         Begin VB.CommandButton cmdEXAM_NEW 
            Caption         =   "&Add"
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
            Left            =   90
            Picture         =   "frmAdd_TYPEofEXAM.frx":2CD8
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Add Type OF Exam"
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Double Click to Edit"
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
         Index           =   4
         Left            =   120
         TabIndex        =   35
         Top             =   2580
         Width           =   1845
      End
   End
   Begin VB.Label lblDEPTID 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   23
      Top             =   5640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblDEPT 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   5970
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmAISAdd_TYPEofEXAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVE_EDIT_NEW_EXAM                                                As String

Function AlternateSAVEandADD(COND As Boolean)
    picSave.Visible = COND
    picAdd.Visible = Not COND
End Function

Function GetExamDescription(EXAMID As Integer) As String
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select ExamDescription From HRMS_ExamType Where ExamID = " & EXAMID & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetExamDescription = Null2String(RSTMP!ExamDescription)
    End If
End Function

Function GenerateNewExamID() As Integer
    Dim RSTMP                                                         As ADODB.Recordset
    Dim ID                                                            As Integer

    GenerateNewExamID = 0
    Set RSTMP = gconDMIS.Execute("Select ExamID From HRMS_EXAMTYPE Order By ExamDescription ASC")

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            ID = RSTMP!DEPT_ID
            RSTMP.MoveNext
        Loop
    End If
    GenerateNewExamID = ID + 1
End Function

Function GenerateNewDeptID() As Integer
    Dim RSTMP                                                         As ADODB.Recordset
    Dim ID                                                            As Integer

    GenerateNewDeptID = 0
    Set RSTMP = gconDMIS.Execute("Select ID From HRMS_DEPARTMENT Order By Dept_ID ASC")

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            ID = RSTMP!DEPT_ID
            RSTMP.MoveNext
        Loop
    End If
    GenerateNewDeptID = ID + 1
End Function

Function CheckIfDepartmentAlreadyExist() As Boolean
    Dim RSTMP                                                         As ADODB.Recordset
    Dim vtxtDEPT                                                      As String

    vtxtDEPT = N2Str2Null(txtDEPT)
    Set RSTMP = gconDMIS.Execute("Select DeptName From HRMS_Department Where DeptName = " & vtxtDEPT & "")

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckIfDepartmentAlreadyExist = True
    Else
        CheckIfDepartmentAlreadyExist = False
    End If
End Function

Function CleanAddExamForm()
    txtID.Text = ""
    txtEXAMDESC.Text = ""
    txtPGRADE.Text = ""
    txtMAX.Text = ""
    txtMIN.Text = ""
End Function

Function GenerateNewIDForExam()
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select ExamID From HRMS_ExamType Order By ExamID ASC")

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            txtID.Text = RSTMP!EXAMID
            RSTMP.MoveNext
        Loop
    End If
    txtID.Text = val(txtID) + 1
    Set RSTMP = Nothing
End Function

Function ValidScoreGiven() As Boolean
    If CDbl(txtMAX) < CDbl(txtPGRADE) Then                    'MAX < PASSING
        MsgBox "Max Score must be Greater than Passing Grade", vbExclamation, "Add Exam"
        On Error Resume Next
        txtMAX.SetFocus
        ValidScoreGiven = False
        Exit Function
    End If
    If CDbl(txtMAX) < CDbl(txtMIN) Then                       'MAX < MIN
        MsgBox "Max Score must be Greater than Min Grade", vbExclamation, "Add Exam"
        On Error Resume Next
        txtMAX.SetFocus
        ValidScoreGiven = False
        Exit Function
    End If

    If CDbl(txtMIN) > CDbl(txtPGRADE) Then                    'MIN > PASSING
        MsgBox "Min Score must be Less than Passing Grade", vbExclamation, "Add Exam"
        On Error Resume Next
        txtMIN.SetFocus
        ValidScoreGiven = False
        Exit Function
    End If
    If CDbl(txtMIN) > CDbl(txtMAX) Then                       'MIN > MAX
        MsgBox "Min Score must be Less than Max Grade", vbExclamation, "Add Exam"
        On Error Resume Next
        txtMIN.SetFocus
        ValidScoreGiven = False
        Exit Function
    End If

    If CDbl(txtPGRADE) < CDbl(txtMIN) Then                    'PASSING < MIN
        MsgBox "Passing Score must be Greater than Min Grade", vbExclamation, "Add Exam"
        On Error Resume Next
        txtPGRADE.SetFocus
        ValidScoreGiven = False
        Exit Function
    End If
    If CDbl(txtPGRADE) > CDbl(txtMAX) Then                    'PASSING > MAX
        MsgBox "Passing Score must be Less than Max Grade", vbExclamation, "Add Exam"
        On Error Resume Next
        txtPGRADE.SetFocus
        ValidScoreGiven = False
        Exit Function
    End If

    ValidScoreGiven = True
End Function

Function CheckTypeOfExamExist() As Boolean
    Dim RSTMP                                                         As ADODB.Recordset
    Dim EXAMDESC                                                      As String

    EXAMDESC = N2Str2Null(txtEXAMDESC)

    Set RSTMP = gconDMIS.Execute("Select ExamDescription,ExamID From HRMS_ExamType Where ExamDescription = " & _
                      EXAMDESC & "")

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If CInt(txtID) = CInt(Null2String(RSTMP!EXAMID)) Then
            CheckTypeOfExamExist = False
        Else
            CheckTypeOfExamExist = True
        End If
    Else
        CheckTypeOfExamExist = False
    End If
End Function

Private Sub cboDEPT_Change()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim ITEM                                                          As ListItem
    Dim EXAM_DESC                                                     As String

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_DEPARTMENT_EXAM Where Department_ID = " & Right(cboDept, 3) & "")
    lsvDEPT_EXAM.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvDEPT_EXAM.ListItems.Add(, , RSTMP!EXAMID)
            EXAM_DESC = GetExamDescription(RSTMP!EXAMID)
            ITEM.SubItems(1) = EXAM_DESC
            ITEM.SubItems(2) = RSTMP!EXAMID

            RSTMP.MoveNext
        Loop
    End If
End Sub

Private Sub cboDEPT_Click()
    Call cboDEPT_Change
End Sub

Private Sub cmdADD_to_DEPT_Click()
    Dim RSTMP                                                         As ADODB.Recordset

    If MsgBox("Save Exam to this Department", vbQuestion + vbYesNo + vbDefaultButton1, "Are You Sure") = vbYes Then
        Set RSTMP = gconDMIS.Execute("Select * From HRMS_DEPARTMENT_EXAM Where Department_ID = " & _
                                     CInt(Right(cboDept, 3)) & " And ExamID = " & CInt(txtID) & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            MsgBox "Exam Type Already Save to this Department", vbInformation, "Department Exam"
        Else
            gconDMIS.Execute ("Insert Into HRMS_DEPARTMENT_EXAM Values(" & CInt(Right(cboDept, 3)) & _
                              "," & CInt(txtID) & ")")
            Call cboDEPT_Change
        End If
    End If
End Sub

Private Sub cmdDEPT_EDIT_Click()
    If cmdDEPT_EDIT.Caption = "EDIT" Then                     'Edit
        lblDEPTID.Caption = Right(cboDept, 3)
        lblDEPT.Caption = "EDIT"

        '        picENTRY.Enabled = False
        '        picEXAM.Enabled = False

        txtDEPT.Text = Mid(cboDept, 1, Len(cboDept) - 6)
        cmdSave.Caption = "UPDATE"

        Call EnabledPicture(False)
        picEDIT_DEPT.Visible = True
        On Error Resume Next
        txtDEPT.SetFocus
    End If
    '    Else                                        'Cancel
    '        cmdDEPT_NEW.Caption = "NEW"
    '        cmdDEPT_EDIT.Caption = "EDIT"
    '
    '        picENTRY.Enabled = True
    '        picEXAM.Enabled = True
    '
    '        cboDEPT.Visible = True
    '        txtDEPT.Visible = False
    '
    '        cboDEPT.SetFocus
    '    End If
End Sub

Private Sub cmdDEPT_NEW_Click()



    txtDEPT.Text = ""
    Call EnabledPicture(False)
    cmdSave.Caption = "SAVE"
    picEDIT_DEPT.Visible = True
    On Error Resume Next
    txtDEPT.SetFocus
End Sub

Private Sub cmdEXAM_CANCEL_Click()
    Call AlternateSAVEandADD(False)
    Call CleanAddExamForm

    picENTRY.Visible = True
    picEDIT_EXAM.Visible = False

    picEDIT_EXAM.Visible = False
    picAdd.Visible = True

    Call EnabledPicture(True)
    cmdEXAM_SAVE.Caption = "Save"
End Sub

Private Sub cmdEXAM_EXIT_Click()
    Unload Me
End Sub

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:14

Private Sub cmdEXAM_NEW_Click()
    If Function_Access(LOGID, "ACESS_ADD", "APPLICANT EXAM TYPE") = False Then Exit Sub
    On Error GoTo Errorcode:

    SAVE_EDIT_NEW_EXAM = "SAVE"
    Call AlternateSAVEandADD(True)
    Call CleanAddExamForm
    Call GenerateNewIDForExam

    cmdEXAM_SAVE.Caption = "Save"
    Call EnabledPicture(False)

    picAdd.Visible = False
    picENTRY.Visible = False
    picEDIT_EXAM.Visible = True

    On Error Resume Next
    txtEXAMDESC.SetFocus

    Exit Sub

Errorcode:
    ShowVBError
End Sub

'Function Feature   :
'Date               : 7/11/2007
'Last Update        : 7/11/2007
'Database Update    :
'Who Updated        : AXP
'Upating Code       : AXP-0707200711:14
Private Sub cmdEXAM_SAVE_Click()
    Dim vtxtID As String, VtxtEXAM_TYPE As String, VtxtPGRADE         As String
    Dim vtxtMAX As String, VtxtMIN                                    As String

    On Error GoTo Errorcode:

    If Not txtEXAMDESC.Text = "" And Not txtPGRADE.Text = "" And _
       Not txtMAX.Text = "" And Not txtMIN.Text = "" And Not txtPGRADE.Text = "" Then

        If Not IsNumeric(txtPGRADE) Then
            MsgBox "Invalid Entry", vbExclamation, "Add Exam Type"
            On Error Resume Next
            txtPGRADE.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtMIN) Then
            MsgBox "Invalid Entry", vbExclamation, "Add Exam Type"
            On Error Resume Next
            txtMIN.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtMAX) Then
            MsgBox "Invalid Entry", vbExclamation, "Add Exam Type"
            On Error Resume Next
            txtMAX.SetFocus
            Exit Sub
        End If

        If ValidScoreGiven Then
            vtxtID = CInt(txtID)
            VtxtEXAM_TYPE = N2Str2Null(txtEXAMDESC)
            VtxtPGRADE = CDbl(txtPGRADE)
            vtxtMAX = CDbl(txtMAX)
            VtxtMIN = CDbl(txtMIN)

            frmMain.MousePointer = 11
            If CheckTypeOfExamExist = False Then
                If SAVE_EDIT_NEW_EXAM = "SAVE" Then
                    gconDMIS.Execute ("Insert Into HRMS_ExamType Values(" & vtxtID & "," & _
                                      VtxtEXAM_TYPE & _
                                      "," & VtxtPGRADE & _
                                      ",'" & "Passed" & _
                                      "'," & VtxtMIN & _
                                      ",'" & "Work Hard" & _
                                      "'," & vtxtMAX & _
                                      ",'" & "Amazing" & "')")

                    cmdADD_to_DEPT.Visible = False
                    Call DisplayAllExamType
                    Call AlternateSAVEandADD(False)
                    Call CleanAddExamForm
                    picEDIT_EXAM.Visible = False
                    'picADD.Visible = True
                    picENTRY.Visible = True

                    Call EnabledPicture(True)
                Else                                          '===================================EDIT
                    gconDMIS.Execute ("Update HRMS_ExamType Set ExamDescription = " & VtxtEXAM_TYPE & _
                                      ",Passing = " & VtxtPGRADE & _
                                      ",MaxScore = " & vtxtMAX & _
                                      ",MinScore = " & VtxtMIN & _
                                    " Where ExamID = " & vtxtID & "")

                    cmdADD_to_DEPT.Visible = False
                    Call DisplayAllExamType
                    Call AlternateSAVEandADD(False)
                    Call CleanAddExamForm

                    picEDIT_EXAM.Visible = False
                    picENTRY.Visible = True

                    Call EnabledPicture(True)
                End If
            Else
                MsgBox "Exam Type Already Exist", vbExclamation, "Add Exam Type"
                On Error Resume Next
                txtEXAMDESC.SetFocus
            End If
        End If
    Else
        MsgBox "Incomplte Entry", vbExclamation, "Type Of Exam"
        txtEXAMDESC.SetFocus
    End If
    frmMain.MousePointer = 11
    frmMain.MousePointer = 0

    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 11
End Sub

Private Sub cmdSave_Click()

    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    Dim vtxtDEPT As String, S_ID                                      As String
    Dim DEPT_ID                                                       As Integer

    If Not txtDEPT.Text = "" Then
        vtxtDEPT = N2Str2Null(txtDEPT)

        If cmdSave.Caption = "SAVE" Then                      'Save
            If CheckIfDepartmentAlreadyExist = True Then
                MsgBox "Department Already Exsit", vbInformation, "Add Department"
                On Error Resume Next
                txtDEPT.SetFocus
                Exit Sub
            Else
                DEPT_ID = GenerateNewDeptID
                S_ID = DEPT_ID

                gconDMIS.Execute ("Insert Into HRMS_DEPARTMENT Values(" & DEPT_ID & "," & vtxtDEPT & ")")

                Call DisplayAllDepartments
                Call EnabledPicture(True)
                picEDIT_DEPT.Visible = False
                On Error Resume Next
                cboDept.SetFocus
            End If
        Else                                                  'Update
            If CheckIfDepartmentAlreadyExist = True Then
                MsgBox "Department Already Exsit", vbInformation, "Add Department"
                On Error Resume Next
                txtDEPT.SetFocus
                Exit Sub
            Else
                DEPT_ID = lblDEPTID.Caption

                gconDMIS.Execute ("Update HRMS_DEPARTMENT Set DeptName = " & vtxtDEPT & _
                                " Where ID = " & DEPT_ID)

                Call DisplayAllDepartments
                Call EnabledPicture(True)
                picEDIT_DEPT.Visible = False
                On Error Resume Next
                cboDept.SetFocus
            End If
        End If
    Else
        MsgBox "Enter a Department name", vbInformation, "New Department"
        txtDEPT.SetFocus
    End If

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cndCANCEL_Click()
    Call EnabledPicture(True)
    picEDIT_DEPT.Visible = False
    On Error Resume Next
    cboDept.SetFocus
End Sub

Private Sub cmdPrint_Click()
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "ACESS_PRINT", "APPLICANT EXAM TYPE") = False Then Exit Sub
    frmMain.MousePointer = 11

    Call PrintSQLReport(rptEXAM, AIS_REPORT_PATH & "ExamType.rpt", "", AIS_REPORT_Connection, 1)

    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    frmMain.MousePointer = 11

    Call DisplayAllExamType
    Call DisplayAllDepartments

    frmMain.MousePointer = 0
End Sub

Private Sub DisplayAllDepartments()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim SZERO                                                         As String

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_DEPARTMENT Order By DeptName ASC")

    cboDept.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            If Len(RSTMP!ID) = 1 Then SZERO = "00"
            If Len(RSTMP!ID) = 2 Then SZERO = "0"

            cboDept.AddItem RSTMP!DEPTNAME & " - " & SZERO & RSTMP!ID
            RSTMP.MoveNext
        Loop
    End If
    cboDept.ListIndex = 0
End Sub

Private Sub DisplayAllExamType()
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_ExamType Order By ExamID")

    lsvEXAM.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvEXAM.ListItems.Add(, , RSTMP!EXAMID)
            ITEM.SubItems(1) = Null2String(RSTMP!ExamDescription)
            ITEM.SubItems(2) = RSTMP!Passing
            ITEM.SubItems(3) = RSTMP!MinScore
            ITEM.SubItems(4) = RSTMP!MaxScore

            RSTMP.MoveNext
        Loop
    End If
End Sub

Private Sub lsvDEPT_EXAM_DblClick()
    Dim Index                                                         As Integer

    If Not lsvDEPT_EXAM.ListItems.count = 0 Then
        Index = CInt(lsvDEPT_EXAM.SelectedItem.Index)
        With lsvDEPT_EXAM
            If MsgBox("Delete Exam Type to This Department", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
                gconDMIS.Execute ("Delete From HRMS_DEPARTMENT_EXAM Where Department_ID = " & _
                                  CInt(Right(cboDept, 3)) & " And ExamID = " & _
                                  CInt(.ListItems(Index).SubItems(2)) & "")
                Call cboDEPT_Change
            End If
        End With
    End If
End Sub

Private Sub lsvDEPT_EXAM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lsvDEPT_EXAM_DblClick
End Sub

Private Sub lsvEXAM_Click()
    On Error Resume Next
    Dim Index                                                         As Integer

    If Not lsvEXAM.ListItems.count = 0 Then

        Index = lsvEXAM.SelectedItem.Index

        With lsvEXAM
            Call DisplayToBeEditExam(.ListItems(Index).Text)
        End With
    End If
End Sub

Private Sub lsvEXAM_DblClick()
    Dim Index                                                         As Integer

    If Not lsvEXAM.ListItems.count = 0 Then
        If Function_Access(LOGID, "ACESS_EDIT", "APPLICANT EXAM TYPE") = False Then Exit Sub
        Index = lsvEXAM.SelectedItem.Index
        SAVE_EDIT_NEW_EXAM = "EDIT"
        picAdd.Visible = False
        picAdd.Visible = False
        picENTRY.Visible = False
        picEDIT_EXAM.Visible = True
        'cmdADD_to_DEPT.Visible = True

        Call AlternateSAVEandADD(True)
        With lsvEXAM
            Call EnabledPicture(False)
            picEDIT_EXAM.Visible = True

            Call DisplayToBeEditExam(.ListItems(Index).Text)
            cmdEXAM_SAVE.Caption = "&UPDATE"
        End With
    End If
End Sub

Private Sub EnabledPicture(COND As Boolean)
    picDEPT.Enabled = COND
    picENTRY.Enabled = COND
    picEXAM.Enabled = COND
End Sub

Private Sub DisplayToBeEditExam(ID As Integer)
    Dim RSTMP                                                         As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_ExamType Where ExamID = " & ID & "")

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        txtID.Text = Null2String(RSTMP!EXAMID)
        txtEXAMDESC.Text = Null2String(RSTMP!ExamDescription)
        txtPGRADE.Text = Null2String(RSTMP!Passing)
        txtMAX.Text = Null2String(RSTMP!MaxScore)
        txtMIN.Text = Null2String(RSTMP!MinScore)
    End If
    Set RSTMP = Nothing
End Sub

Private Sub lsvEXAM_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    Call lsvEXAM_Click
End Sub

Private Sub lsvEXAM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lsvEXAM_DblClick
End Sub

