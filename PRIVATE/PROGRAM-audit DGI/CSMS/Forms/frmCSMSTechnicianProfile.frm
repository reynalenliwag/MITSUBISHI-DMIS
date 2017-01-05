VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmCSMSTechnicianProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Technician Profile"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   Icon            =   "frmCSMSTechnicianProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameTech 
      Caption         =   "Technician Profile"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7665
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   7305
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   0
            Left            =   3030
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   1380
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   1
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   570
            Width           =   3435
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   2
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   900
            Width           =   3420
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   3
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "First Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   21
            Top             =   510
            Width           =   840
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Employee No:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1950
            TabIndex        =   20
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   780
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Code:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   480
            TabIndex        =   18
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   510
         Left            =   6690
         TabIndex        =   12
         Top             =   6720
         Width           =   840
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   4815
         Left            =   60
         TabIndex        =   13
         Top             =   1770
         Width           =   7455
         _Version        =   655364
         _ExtentX        =   13150
         _ExtentY        =   8493
         _StockProps     =   64
         Appearance      =   9
         Color           =   4
         PaintManager.BoldSelected=   -1  'True
         PaintManager.OneNoteColors=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         PaintManager.FixedTabWidth=   120
         PaintManager.MinTabWidth=   110
         ItemCount       =   3
         Item(0).Caption =   "Skill"
         Item(0).Tooltip =   "Technician Skill"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "Picture1"
         Item(1).Caption =   "Exam"
         Item(1).Tooltip =   "Technician Exam Taken"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "Picture2"
         Item(2).Caption =   "Training"
         Item(2).Tooltip =   "Technician Training Attended"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "Picture3"
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   4215
            Left            =   -70000
            ScaleHeight     =   4185
            ScaleWidth      =   7395
            TabIndex        =   27
            Top             =   600
            Visible         =   0   'False
            Width           =   7425
            Begin MSComctlLib.ListView listtraining 
               Height          =   4155
               Left            =   0
               TabIndex        =   28
               Top             =   0
               Width           =   7380
               _ExtentX        =   13018
               _ExtentY        =   7329
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   5
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "No"
                  Object.Width           =   706
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Training"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Monyear"
                  Object.Width           =   1411
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Place"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Sponsor"
                  Object.Width           =   3175
               EndProperty
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   4215
            Left            =   -70000
            ScaleHeight     =   4185
            ScaleWidth      =   7395
            TabIndex        =   25
            Top             =   600
            Visible         =   0   'False
            Width           =   7425
            Begin MSComctlLib.ListView ListExam 
               Height          =   4170
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Width           =   7395
               _ExtentX        =   13044
               _ExtentY        =   7355
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
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   6
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "No"
                  Object.Width           =   706
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Examination"
                  Object.Width           =   4304
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "DateTaken"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Rating"
                  Object.Width           =   1411
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Place"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "LastUpdate"
                  Object.Width           =   1764
               EndProperty
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   4215
            Left            =   0
            ScaleHeight     =   4185
            ScaleWidth      =   7395
            TabIndex        =   23
            Top             =   600
            Width           =   7425
            Begin MSComctlLib.ListView ListSkill 
               Height          =   4185
               Left            =   0
               TabIndex        =   24
               Top             =   0
               Width           =   7395
               _ExtentX        =   13044
               _ExtentY        =   7382
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
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "No"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Skill"
                  Object.Width           =   10407
               EndProperty
            End
         End
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   6690
      TabIndex        =   10
      Top             =   6810
      Width           =   945
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   0
      TabIndex        =   8
      Top             =   1530
      Width           =   7710
      Begin MSComctlLib.ListView ListResult 
         Height          =   4800
         Left            =   60
         TabIndex        =   9
         ToolTipText     =   "Double Click To View Profile"
         Top             =   255
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   8467
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
         Appearance      =   1
         MousePointer    =   99
         MouseIcon       =   "frmCSMSTechnicianProfile.frx":0E42
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NO"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp N0"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Technician"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Comple Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "First Name "
            Object.Width           =   4304
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Criteria"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   30
      TabIndex        =   1
      Top             =   300
      Width           =   7635
      Begin VB.TextBox txtkeyword 
         Height          =   360
         Left            =   1290
         TabIndex        =   7
         Top             =   660
         Width           =   4095
      End
      Begin VB.OptionButton Otp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Employee No"
         Height          =   270
         Index           =   3
         Left            =   4200
         TabIndex        =   5
         Top             =   315
         Width           =   1425
      End
      Begin VB.OptionButton Otp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Comple Name"
         Height          =   270
         Index           =   2
         Left            =   2700
         TabIndex        =   4
         Top             =   300
         Width           =   1425
      End
      Begin VB.OptionButton Otp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "First Name"
         Height          =   315
         Index           =   1
         Left            =   1545
         TabIndex        =   3
         Top             =   285
         Width           =   1230
      End
      Begin VB.OptionButton Otp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Technician"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   2
         Top             =   255
         Width           =   1230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Keyword"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   285
         TabIndex        =   6
         Top             =   705
         Width           =   870
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seacrh Technician"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   390
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   2520
   End
End
Attribute VB_Name = "frmCSMSTechnicianProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TheEmpNO                                           As String

Sub DisplayResult()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim Keyword                                        As String
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer

    SQL = "SELECT * FROM CSMS_vw_Technician WHERE "
    Keyword = Trim(txtkeyword.Text)

    If Len(txtkeyword.Text) = 0 Then Exit Sub

    If Otp(0).Value = True Then
        SQL = SQL & "technician LIKE '" & Keyword & "%'"
    End If

    If Otp(1).Value = True Then
        SQL = SQL & "firstname LIKE '" & Keyword & "%'"
    End If

    If Otp(2).Value = True Then
        SQL = SQL & " tech_name LIKE '" & Keyword & "%'"
    End If

    If Otp(3).Value = True Then
        SQL = SQL & " empno LIKE '" & Keyword & "%'"
    End If

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    cnt = 0
    ListResult.ListItems.Clear

    Do While Not RS.EOF
        cnt = cnt + 1
        Set ITEM = ListResult.ListItems.Add(, , cnt)
        ITEM.SubItems(1) = Null2String(RS!EMPNO)
        ITEM.SubItems(2) = Null2String(RS!Technician)
        ITEM.SubItems(3) = Null2String(RS!TECH_NAME)
        ITEM.SubItems(4) = Null2String(RS!Firstname)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub displayExam()
    Dim SQL                                            As String
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer
    Dim RS                                             As New ADODB.Recordset
    Dim theEmpCode                                     As String

    theEmpCode = TheEmpNO

    SQL = "SELECT * FROM HRMS_ExamsPassed where empno='" & theEmpCode & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListExam.ListItems.Clear
    cnt = 0
    Do While Not RS.EOF
        cnt = cnt + 1
        Set ITEM = ListExam.ListItems.Add(, , cnt)
        ITEM.SubItems(1) = Null2String(RS!examination)
        ITEM.SubItems(2) = Null2String(RS!dateTaken)
        ITEM.SubItems(3) = Null2String(RS!Rating)
        ITEM.SubItems(4) = Null2String(RS!place)
        ITEM.SubItems(5) = Null2String(RS!lastupdate)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub displaySkill()
    Dim SQL                                            As String
    Dim ITEM                                           As ListItem
    Dim RS                                             As New ADODB.Recordset
    Dim cnt                                            As Integer
    Dim thecode                                        As String

    thecode = TheEmpNO

    SQL = "SELECT * FROM CSMS_vw_TechSkills where empno='" & thecode & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    cnt = 0
    ListSkill.ListItems.Clear
    Do While Not RS.EOF
        cnt = cnt + 1
        Set ITEM = ListSkill.ListItems.Add(, , cnt)
        ITEM.SubItems(1) = Null2String(RS!SKILLS)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub displayTraining()
    Dim SQL                                            As String
    Dim cnt                                            As Integer
    Dim RS                                             As New ADODB.Recordset
    Dim thecode                                        As String
    Dim ITEM                                           As ListItem

    thecode = TheEmpNO

    SQL = "SELECT * FROM HRMS_Training where empno='" & thecode & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listtraining.ListItems.Clear
    cnt = 0
    With RS
        Do While Not RS.EOF
            cnt = cnt + 1
            Set ITEM = listtraining.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!training)
            ITEM.SubItems(2) = Null2String(!monyear)
            ITEM.SubItems(3) = Null2String(!place)
            ITEM.SubItems(4) = Null2String(!sponsor)
            RS.MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub displaydefault()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer

    SQL = "SELECT * FROM CSMS_vw_Technician "

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListResult.ListItems.Clear

    cnt = 0

    Do While Not RS.EOF
        cnt = cnt + 1
        Set ITEM = ListResult.ListItems.Add(, , cnt)
        ITEM.SubItems(1) = Null2String(RS!EMPNO)
        ITEM.SubItems(2) = Null2String(RS!Technician)
        ITEM.SubItems(3) = Null2String(RS!TECH_NAME)
        ITEM.SubItems(4) = Null2String(RS!Firstname)
        RS.MoveNext
    Loop
    Set RS = Nothing

End Sub

Private Sub cmdBack_Click()
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    FrameTech.Visible = False
End Sub

Private Sub cmdViewProFile_Click()
    FrameTech.Visible = True
End Sub

Private Sub Command1_Click()
    FrameTech.Visible = False
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    displaydefault

    FrameTech.Visible = False

End Sub

Private Sub ListResult_Click()
    On Error Resume Next
    TheEmpNO = ListResult.ListItems.ITEM(ListResult.SelectedItem.Index)
    TheEmpNO = ListResult.SelectedItem.SubItems(1)


End Sub

Private Sub ListResult_DblClick()
    On Error Resume Next

    TheEmpNO = ListResult.ListItems.ITEM(ListResult.SelectedItem.Index)
    TheEmpNO = ListResult.SelectedItem.SubItems(1)

    txt(0).Text = ListResult.SelectedItem.SubItems(1)
    txt(1).Text = ListResult.SelectedItem.SubItems(4)
    txt(2).Text = ListResult.SelectedItem.SubItems(3)
    txt(3).Text = ListResult.SelectedItem.SubItems(2)


    FrameTech.Visible = True
    Call displayExam
    Call displaySkill
    Call displayTraining

End Sub

Private Sub txtKeyword_Change()
    DisplayResult
End Sub

