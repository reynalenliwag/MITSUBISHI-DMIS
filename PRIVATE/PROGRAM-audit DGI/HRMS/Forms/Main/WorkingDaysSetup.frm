VERSION 5.00
Begin VB.Form frmHRMSWorkingDaysSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Working Days Setup"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4350
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
      Left            =   3510
      MouseIcon       =   "WorkingDaysSetup.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "WorkingDaysSetup.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Exit Window"
      Top             =   990
      Width           =   705
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   2820
      MouseIcon       =   "WorkingDaysSetup.frx":04B8
      MousePointer    =   99  'Custom
      Picture         =   "WorkingDaysSetup.frx":060A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save this Record"
      Top             =   990
      Width           =   705
   End
   Begin VB.TextBox txtWorkInMonth 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   510
      Width           =   735
   End
   Begin VB.TextBox txtWorkYear 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   90
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Average Working days in a Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   570
      Width           =   3345
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Working Days in a Year "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   3345
   End
End
Attribute VB_Name = "frmHRMSWorkingDaysSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Screen.MousePointer = 11
SaveWorkingSetup
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
CenterMe frmMain, Me, 0
StoreWorkingDaysSetup
End Sub

Sub StoreWorkingDaysSetup()
Dim rsWorkingDays As ADODB.Recordset
Set rsWorkingDays = New ADODB.Recordset
Set rsWorkingDays = gconDMIS.Execute("Select * from HRMS_WorkingDaysSetup")
If Not rsWorkingDays.EOF And Not rsWorkingDays.BOF Then
   txtWorkYear.Text = N2Str2Zero(rsWorkingDays!NumberofDaysInYear)
   txtWorkInMonth.Text = N2Str2Zero(rsWorkingDays!AverageinMonth)
End If
End Sub

Sub SaveWorkingSetup()
txtWorkYear.Text = NumericVal(txtWorkYear.Text)
txtWorkInMonth.Text = NumericVal(txtWorkInMonth.Text)
If txtWorkYear.Text > 365 Then
   MsgBox "Number of Days in Year is Invalid!", vbCritical, "Invalid Input"
   Exit Sub
End If

If txtWorkInMonth.Text > 31 Then
   MsgBox "Average Days in Month is Invalid!", vbCritical, "Invalid Input"
   Exit Sub
End If

gconDMIS.Execute ("Update HRMS_WorkingDaysSetup Set " & _
                 "NumberofDaysInYear = " & txtWorkYear.Text & "," & _
                 "AverageinMonth = " & txtWorkInMonth.Text)
ShowSuccessFullyUpdated
Unload Me
End Sub
