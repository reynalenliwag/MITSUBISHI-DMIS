VERSION 5.00
Begin VB.Form frmHRMS_SelectCutOff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cut-Off Selection"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   3510
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   900
      MouseIcon       =   "SelectCutOff.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "SelectCutOff.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Edit Selected Record"
      Top             =   2520
      Width           =   705
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1620
      MouseIcon       =   "SelectCutOff.frx":04AE
      MousePointer    =   99  'Custom
      Picture         =   "SelectCutOff.frx":0600
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exit Window"
      Top             =   2520
      Width           =   705
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   3195
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1290
      Width           =   3195
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   3195
   End
   Begin VB.Label Label3 
      Caption         =   "Select Year"
      Height          =   315
      Left            =   690
      TabIndex        =   5
      Top             =   1680
      Width           =   1875
   End
   Begin VB.Label Label2 
      Caption         =   "Select Month"
      Height          =   315
      Left            =   690
      TabIndex        =   4
      Top             =   930
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Select Cut-off"
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Top             =   180
      Width           =   1875
   End
End
Attribute VB_Name = "frmHRMS_SelectCutOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    If Combo1.Text = "1st Cut-Off" Then
        CUTTOFF_CODE = "1"
        CUTTOFF_CODE_STR = "1st Cut-Off"
    ElseIf Combo1.Text = "2nd Cut-Off" Then
        CUTTOFF_CODE = "2"
        CUTTOFF_CODE_STR = "2nd Cut-Off"
    ElseIf Combo1.Text = "" Then
        MsgBox "Cut-Off Code...This field is empty"
        Exit Sub
    End If
    PAY_MONTH = What_month(Combo2.Text)
    PAY_MONTH_STR = "" & Combo2.Text & ""
    PAY_YEAR = NumericVal(Combo3.Text)
    PAY_YEAR_STR = "" & Combo3.Text & ""
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    FillCombo
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Sub FillCombo()
    Combo1.AddItem "1st Cut-Off"
    Combo1.AddItem "2nd Cut-Off"
    Combo1.AddItem "Monthly Cut-Off"
    
    Combo2.AddItem "January"
    Combo2.AddItem "February"
    Combo2.AddItem "March"
    Combo2.AddItem "April"
    Combo2.AddItem "May"
    Combo2.AddItem "June"
    Combo2.AddItem "July"
    Combo2.AddItem "August"
    Combo2.AddItem "September"
    Combo2.AddItem "October"
    Combo2.AddItem "November"
    Combo2.AddItem "December"
    
    
    Combo3.AddItem "" & Year(DateAdd("yyyy", -3, Date)) & ""
    Combo3.AddItem "" & Year(DateAdd("yyyy", -2, Date)) & ""
    Combo3.AddItem "" & Year(DateAdd("yyyy", -1, Date)) & ""
    Combo3.AddItem "" & Year(Date) & ""
    Combo3.AddItem "" & Year(DateAdd("yyyy", 1, Date)) & ""
    Combo3.AddItem "" & Year(DateAdd("yyyy", 2, Date)) & ""
    Combo3.AddItem "" & Year(DateAdd("yyyy", 3, Date)) & ""
    
End Sub

