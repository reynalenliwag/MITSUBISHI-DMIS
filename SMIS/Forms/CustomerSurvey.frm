VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmSMIS_Files_Survey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Survey Data Entry"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CustomerSurvey.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   10020
   Begin VB.CommandButton Command7 
      Caption         =   "Edit"
      Height          =   405
      Left            =   4560
      TabIndex        =   43
      Top             =   2640
      Width           =   915
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Del"
      Height          =   405
      Left            =   5520
      TabIndex        =   42
      Top             =   2640
      Width           =   915
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add"
      Height          =   405
      Left            =   3600
      TabIndex        =   41
      Top             =   2640
      Width           =   915
   End
   Begin VB.PictureBox picOE 
      BackColor       =   &H00000080&
      Height          =   4365
      Left            =   6510
      ScaleHeight     =   4305
      ScaleWidth      =   3375
      TabIndex        =   3
      Top             =   2610
      Width           =   3435
      Begin VB.TextBox txtOEValue 
         Height          =   2775
         Left            =   240
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   1380
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Best Answers"
         Height          =   465
         Left            =   300
         TabIndex        =   39
         Top             =   810
         Width           =   2055
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   630
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Question"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1590
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   180
      Width           =   3915
   End
   Begin VB.PictureBox picNU 
      BackColor       =   &H008080FF&
      Height          =   4365
      Left            =   6510
      ScaleHeight     =   4305
      ScaleWidth      =   3405
      TabIndex        =   7
      Top             =   2610
      Width           =   3465
      Begin VB.TextBox txtNUValue 
         Height          =   585
         Left            =   480
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   1350
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Average Value"
         Height          =   375
         Left            =   540
         TabIndex        =   37
         Top             =   990
         Width           =   1425
      End
   End
   Begin VB.PictureBox picMC 
      BackColor       =   &H0080FF80&
      Height          =   4185
      Left            =   6510
      ScaleHeight     =   4125
      ScaleWidth      =   3315
      TabIndex        =   5
      Top             =   2610
      Width           =   3375
      Begin VB.OptionButton optMC 
         Caption         =   "Option1"
         Height          =   345
         Index           =   5
         Left            =   330
         TabIndex        =   19
         Top             =   2280
         Width           =   2235
      End
      Begin VB.OptionButton optMC 
         Caption         =   "Option1"
         Height          =   345
         Index           =   4
         Left            =   330
         TabIndex        =   18
         Top             =   1860
         Width           =   2235
      End
      Begin VB.OptionButton optMC 
         Caption         =   "Option1"
         Height          =   345
         Index           =   3
         Left            =   330
         TabIndex        =   17
         Top             =   1440
         Width           =   2235
      End
      Begin VB.OptionButton optMC 
         Caption         =   "Option1"
         Height          =   345
         Index           =   2
         Left            =   330
         TabIndex        =   16
         Top             =   1050
         Width           =   2235
      End
      Begin VB.OptionButton optMC 
         Caption         =   "Option1"
         Height          =   345
         Index           =   1
         Left            =   330
         TabIndex        =   15
         Top             =   660
         Width           =   2235
      End
      Begin VB.OptionButton optMC 
         Caption         =   "Option1"
         Height          =   345
         Index           =   0
         Left            =   330
         TabIndex        =   14
         Top             =   240
         Width           =   2235
      End
   End
   Begin VB.PictureBox picOR 
      BackColor       =   &H00FF0000&
      Height          =   4245
      Left            =   6510
      ScaleHeight     =   4185
      ScaleWidth      =   3315
      TabIndex        =   4
      Top             =   2610
      Width           =   3375
      Begin VB.CommandButton Command4 
         Caption         =   "Update"
         Height          =   435
         Left            =   1200
         TabIndex        =   33
         Top             =   3450
         Width           =   915
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   2160
         TabIndex        =   32
         Top             =   3450
         Width           =   915
      End
      Begin VB.ComboBox cboOR 
         Height          =   345
         Index           =   5
         Left            =   90
         TabIndex        =   31
         Text            =   "Combo2"
         Top             =   2370
         Width           =   705
      End
      Begin VB.TextBox txtOR 
         Height          =   345
         Index           =   5
         Left            =   810
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   2370
         Width           =   2475
      End
      Begin VB.ComboBox cboOR 
         Height          =   345
         Index           =   4
         Left            =   60
         TabIndex        =   29
         Text            =   "Combo2"
         Top             =   1980
         Width           =   705
      End
      Begin VB.TextBox txtOR 
         Height          =   345
         Index           =   4
         Left            =   780
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1980
         Width           =   2475
      End
      Begin VB.ComboBox cboOR 
         Height          =   345
         Index           =   3
         Left            =   90
         TabIndex        =   27
         Text            =   "Combo2"
         Top             =   1590
         Width           =   705
      End
      Begin VB.TextBox txtOR 
         Height          =   345
         Index           =   3
         Left            =   810
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1590
         Width           =   2475
      End
      Begin VB.ComboBox cboOR 
         Height          =   345
         Index           =   2
         Left            =   60
         TabIndex        =   25
         Text            =   "Combo2"
         Top             =   1200
         Width           =   705
      End
      Begin VB.TextBox txtOR 
         Height          =   345
         Index           =   2
         Left            =   780
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1200
         Width           =   2475
      End
      Begin VB.ComboBox cboOR 
         Height          =   345
         Index           =   1
         Left            =   60
         TabIndex        =   23
         Text            =   "Combo2"
         Top             =   810
         Width           =   705
      End
      Begin VB.TextBox txtOR 
         Height          =   345
         Index           =   1
         Left            =   780
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   810
         Width           =   2475
      End
      Begin VB.ComboBox cboOR 
         Height          =   345
         Index           =   0
         Left            =   30
         TabIndex        =   21
         Text            =   "Combo2"
         Top             =   420
         Width           =   705
      End
      Begin VB.TextBox txtOR 
         Height          =   345
         Index           =   0
         Left            =   750
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   420
         Width           =   2475
      End
   End
   Begin VB.PictureBox picLS 
      BackColor       =   &H0080C0FF&
      Height          =   4185
      Left            =   6510
      ScaleHeight     =   4125
      ScaleWidth      =   3315
      TabIndex        =   6
      Top             =   2610
      Width           =   3375
      Begin VB.TextBox txtLSMinRating 
         Height          =   345
         Left            =   0
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtLSMaxRating 
         Height          =   345
         Left            =   60
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   300
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   1710
         TabIndex        =   13
         Top             =   3660
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update"
         Height          =   435
         Left            =   750
         TabIndex        =   12
         Top             =   3660
         Width           =   915
      End
      Begin VB.ComboBox cboLSRating 
         Height          =   345
         Left            =   930
         TabIndex        =   11
         Text            =   "Combo2"
         Top             =   1770
         Width           =   1845
      End
      Begin VB.TextBox txtMaxQuestion 
         Height          =   345
         Left            =   810
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   870
         Width           =   1935
      End
      Begin VB.TextBox txtMinQuestion 
         Height          =   345
         Left            =   810
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   270
         Width           =   1995
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Scale"
         Height          =   225
         Left            =   30
         TabIndex        =   8
         Top             =   60
         Width           =   465
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3855
      Left            =   150
      TabIndex        =   40
      Top             =   3090
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Population"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Survey Name"
      Height          =   225
      Left            =   360
      TabIndex        =   0
      Top             =   210
      Width           =   1095
   End
End
Attribute VB_Name = "frmSMIS_Files_Survey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Stype                              As String
Dim SID                                As Long
Dim ID                                 As Long

'Likert-scale: Whenty you want to know respondents'
'feelings or attitudes about something, '
'consider asking a Likert-scale question. '
'The respondents must indicate how closely their feelings '
'match the question or statement on a rating scale. '
'The number at one end of the scale represents least agreement, '
'or "Strongly Disagree," and the number at the other
'end of the scale represents most agreement, or "Strongly Agree." '
'If the scale includes other words at either end to further clarify '
'the meaning of the numbers, it is known as a Likert-style question.



'Multiple-choice: When you want respondents to pick the best answer
'or answers from among all the possible options, consider writing a
'multiple-choice question. Multiple-choice questions are
'easy to lay out on a written survey. '
'Include specific directions about how many answers to select '
'directly after the question. Example:
'
'Why don 't you use the school's cafeteria services? (circle one):
'a.
' It 's too expensive.
'b.
' Serving times conflict with my class schedule.
'c.
' The location is inconvenient.
'd.
' The food quality is poor.
'e.
' Other (please explain):_______________



'Ordinal: When you need all possible answers to be rank ordered, ask an ordinal question. Example:
'
'Please write a number between 1 and 5 next to each item below. Put a 1 next to the item that is MOST important to you in selecting an on-line university course. Put a 5 next to the item that is LEAST important. Please use each number only ONCE.
'
'___ a.
' Availability of instructor for assistance.
'
'___ b.
' Tuition cost for the course.
'
'___ c.
' Ability to work in groups with other students.
'
'___ d.
' Quality and quantity of instructor feedback.
'
'___ e.
' Number of students enrolled.



'Numerical: When the answer must be a real number, ask a
'numerical question. Example: How old were you on your last birthday?

Dim RsSurvey                           As adodb.Recordset

Private Sub Combo1_Change()
    If Combo1.ListIndex = -1 Then: Exit Sub

    FillSurvey Combo1.ItemData(Combo1.ListIndex)
End Sub

Private Sub Combo1_Click()
    Combo1_Change
End Sub

Private Sub Form_Load()
    Set RsSurvey = New adodb.Recordset
    AddColumnHeader "SN, Questions,R Population", ListView1
    ResizeColumnHeader ListView1, "10,69,20"
    FillCombo "Select SName , SID from Cris_SQ_Hdr", 1, 0, Combo1

End Sub
Sub FillSurvey(xxx)
    flex_FillListView gconDMIS.Execute("Select QName, SType, SID,ID from Cris_SQ_Det Where SID=" & xxx), ListView1, True, False
End Sub

Private Sub LISTVIEW1_ItemClick(ByVal Item As MSComctlLib.ListItem)
'QName, SType, SID,ID

    Stype = Item.ListSubItems(2).Text
    SID = CLng(Item.ListSubItems(3).Text)
    ID = CLng(Item.ListSubItems(4).Text)
    picOE.Visible = False
    picNU.Visible = False
    picOR.Visible = False
    picMC.Visible = False
    picLS.Visible = False
    'Likert-scale:LS
    'Multiple-choice:MC
    'Ordinal:OR
    'Numerical:NU
    Me.Caption = Stype
    Select Case Stype
        Case "LS"                                     'Likert-scale
            Fill_LSQuestion ID
            Fill_LSAnswer ID
        Case "MC"                                     'MULTIPLE
            Fill_MCQuestion ID
        Case "OR"                                     ' Ordinal
            Fill_ORQuestion ID
            
        Case "NU"                                     ' NUMERICAL
             picNU.Visible = True
            txtNUValue.Text = 0
            Fill_NUAnswer
        Case "OE"                                     ' NUMERICAL
            picOE.Visible = True
            txtOEValue.Text = 0
            Fill_OEAnswer
    End Select

End Sub
Sub Fill_OEAnswer()
    Dim TempRs                         As adodb.Recordset
    Set TempRs = gconDMIS.Execute("Select * from Cris_SQ_Det WHERE ID=" & ID)
    If Not (TempRs.EOF Or TempRs.BOF) Then
    txtOEValue = TempRs!a1
    End If
End Sub
Sub Fill_NUAnswer()
    Dim TempRs                         As adodb.Recordset
    Set TempRs = gconDMIS.Execute("Select * from Cris_SQ_Det WHERE ID=" & ID)
    If Not (TempRs.EOF Or TempRs.BOF) Then
    txtNUValue = TempRs!a1
    End If
End Sub
Sub Fill_MCQuestion(ID)
    Dim TempRs                         As adodb.Recordset
    picMC.Visible = True
    Set TempRs = gconDMIS.Execute("Select * from Cris_SQ_Det WHERE ID=" & ID)
    If Not (TempRs.EOF Or TempRs.BOF) Then
        'Invisible Visible
        For i = 0 To 5
            optMC(i).Visible = False
        Next
        For i = 0 To TempRs!Freq - 1
            optMC(i).Visible = True
        Next
        optMC(0).Caption = TempRs!S1
        optMC(1).Caption = TempRs!S2
        optMC(2).Caption = TempRs!S3
        optMC(3).Caption = TempRs!S4
        optMC(4).Caption = TempRs!S5
        optMC(5).Caption = TempRs!S6
    End If
End Sub

Sub Fill_LSQuestion(ID)
    Dim TempRs                         As adodb.Recordset
    picLS.Visible = True
    Set TempRs = gconDMIS.Execute("Select * from Cris_SQ_Det WHERE ID=" & ID)
    If Not (TempRs.EOF Or TempRs.BOF) Then
        'Invisible Visible
        'Clear Combo Visible
        ClearCombo picLS
        'Clear Control Visible
        ClearControl picLS
        For i = TempRs!MinScale To TempRs!MaxScale
            cboLSRating.AddItem i
        Next
        txtMinQuestion = TempRs!S1
        txtMaxQuestion = TempRs!S6
        
        txtLSMaxRating = TempRs!MinScale
        txtLSMinRating = TempRs!MaxScale
    End If
End Sub
Sub Fill_LSAnswer(QID)
    Dim TempRs                         As adodb.Recordset
    Set TempRs = gconDMIS.Execute("Select * from Cris_SData_det WHERE QID=" & QID)
    If Not (TempRs.EOF Or TempRs.BOF) Then
        cboLSRating.Text = TempRs.Fields("A1").Value
    End If
End Sub

Sub Fill_ORQuestion(ID)
    Dim TempRs                         As adodb.Recordset
    picOR.Visible = True
    Set TempRs = gconDMIS.Execute("Select * from Cris_SQ_Det WHERE ID=" & ID)
    If Not (TempRs.EOF Or TempRs.BOF) Then
        'Invisible Visible
        For i = 0 To 5
            cboOR(i).Visible = False
            txtOR(i).Visible = False
        Next
        For i = 0 To TempRs!Freq - 1
            cboOR(i).Visible = True
            txtOR(i).Visible = True
        Next
        'Clear Combo Visible
        ClearCombo picOR
        'Clear Control Visible
        ClearControl picOR
        For i = TempRs!MinScale To TempRs!MaxScale
            cboOR(0).AddItem i
            cboOR(1).AddItem i
            cboOR(2).AddItem i
            cboOR(3).AddItem i
            cboOR(4).AddItem i
            cboOR(5).AddItem i
        Next
        txtOR(0) = TempRs!S1
        txtOR(1) = TempRs!S2
        txtOR(2) = TempRs!S3
        txtOR(3) = TempRs!S4
        txtOR(4) = TempRs!S5
        txtOR(5) = TempRs!S6
    End If
End Sub
Sub ClearControl(cControl As Object)
    Dim cntrl                          As Control

    For Each cntrl In Me.ControlS
        If cntrl.Container.hwnd = cControl.hwnd Then
            If TypeOf cntrl Is ComboBox Or TypeOf cntrl Is TextBox Then
                cntrl.Text = ""
            End If
        End If
    Next
End Sub
Sub ClearCombo(cControl As Object)
    Dim cntrl                          As Control
    For Each cntrl In Me.ControlS
        If cntrl.Container.hwnd = cControl.hwnd Then
            If TypeOf cntrl Is ComboBox Then
                cntrl.Clear
            End If
        End If
    Next
End Sub


