VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewCards 
   Caption         =   "Personnel Attendance Tracking System"
   ClientHeight    =   7245
   ClientLeft      =   1365
   ClientTop       =   1095
   ClientWidth     =   9240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7245
   ScaleWidth      =   9240
   Begin VB.ComboBox cboMonth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7080
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      Caption         =   "Month"
      Height          =   990
      Left            =   6975
      TabIndex        =   11
      Top             =   90
      Width           =   2085
      Begin VB.TextBox TxtMonth 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1725
         TabIndex        =   6
         Top             =   270
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Days: 16 - 31"
      Height          =   5115
      Left            =   4680
      TabIndex        =   10
      Top             =   1200
      Width           =   4380
      Begin MSFlexGridLib.MSFlexGrid Grid2 
         Height          =   4815
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8493
         _Version        =   393216
         Rows            =   17
         Cols            =   5
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   150
      TabIndex        =   9
      Top             =   6315
      Width           =   8955
      Begin VB.CommandButton CmdQuit 
         Caption         =   "Back to Attendace Entry"
         Height          =   375
         Left            =   6660
         TabIndex        =   1
         Top             =   210
         Width           =   2145
      End
      Begin VB.CommandButton CmdPrintAll 
         Caption         =   "Print all cards"
         Height          =   375
         Left            =   3315
         TabIndex        =   3
         Top             =   240
         Width           =   2265
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print this card only"
         Height          =   375
         Left            =   150
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Days: 1 - 15"
      Height          =   5115
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   4395
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   4815
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8493
         _Version        =   393216
         Rows            =   17
         Cols            =   5
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Employee's Name"
      Height          =   990
      Left            =   1845
      TabIndex        =   7
      Top             =   90
      Width           =   5070
      Begin VB.TextBox TxtEmpName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   4875
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Emlpoyee Number"
      Height          =   960
      Left            =   135
      TabIndex        =   5
      Top             =   120
      Width           =   1635
      Begin VB.TextBox TxtEmpNumber 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   600
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmViewCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tdate As Date
Dim percas As String
Dim rsEmpInfoCards As ADODB.Recordset

Sub InitGrid()
Dim X, y As Integer
With Grid1
     .ColWidth(0) = 765
     For X = 1 To 4
         .ColWidth(X) = 765
     Next X
     For y = 0 To 16
         .RowHeight(y) = 250
     Next y
     .Row = 0
     .Col = 0
     .Text = "DATE"
     
     .Col = 1
     .Text = "IN-AM"
     
     .Col = 2
     .Text = "OUT-AM"
   
     .Col = 3
     .Text = "IN-PM"
     
     .Col = 4
     .Text = "OUT-PM"
    
     For X = 0 To 4
         .Col = X
         For y = 1 To 16
             .Row = y
             .Text = ""
         Next y
     Next X
End With

With Grid2
     .ColWidth(0) = 765
     For X = 1 To 4
         .ColWidth(X) = 765
     Next X
        
     For y = 0 To 16
         .RowHeight(y) = 250
     Next y
        
     .Row = 0
     .Col = 0
     .Text = "DATE"
   
     .Col = 1
     .Text = "IN-AM"
   
     .Col = 2
     .Text = "OUT-AM"
 
     .Col = 3
     .Text = "IN-PM"
   
     .Col = 4
     .Text = "OUT-PM"

     For X = 0 To 4
         .Col = X
         For y = 1 To 16
             .Row = y
             .Text = ""
         Next y
     Next X
End With
End Sub

Private Sub cboMonth_Change()
InitGrid
FillAll
End Sub

Private Sub cboMonth_Click()
InitGrid
FillAll
End Sub

Private Sub cmdPrint_Click()
If TxtEmpNumber = "" Then
   MsgBox ("Sorry: Can not print this card, employee number not yet selected")
   TxtEmpNumber.SetFocus
Else
   frmPrintCards.Show vbModal
End If
End Sub

'Private Sub CmdPrintAll_Click()
'frmPrintALL.Show vbModal
'End Sub

Private Sub CmdQuit_Click()
Unload Me
frmLOGIN.TxtEmpNumber.SetFocus
End Sub

Private Sub Form_Load()
fillcbomonth
InitGrid
TxtMonth.Text = Date2Month(Date)
cboMonth.Text = Date2Month(Date)
End Sub

Private Sub TxtEmpNumber_GotFocus()
'TxtMonth.Text = Date2Month(Date)
TxtEmpNumber.Text = ""
TxtEmpName.Text = ""
InitGrid
End Sub

Private Sub TxtEmpNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub TxtEmpNumber_LostFocus()
FillAll
End Sub

Sub fillcbomonth()
Dim X As Integer
Dim Thedate As Date
Thedate = OneMonth(Date, -3)
For X = 1 To 4
    Thedate = OneMonth(Thedate, 1)
    cboMonth.AddItem Date2Month(CDate(Thedate))
Next
End Sub

Sub FillAll()
'On Error Resume Next
Dim Criteria As String
Dim k As Integer
Dim r As Integer

If cboMonth.Text = Date2Month(OneMonth(Date, -2)) Then
   tdate = OneMonth(Date, -2)
ElseIf cboMonth.Text = Date2Month(OneMonth(Date, -1)) Then
   tdate = OneMonth(Date, -1)
ElseIf cboMonth.Text = Date2Month(Date) Then
   tdate = Date
ElseIf cboMonth.Text = Date2Month(OneMonth(Date, 1)) Then
   tdate = OneMonth(Date, 1)
End If
        
If TxtEmpNumber.Text = "" Then Exit Sub

If Val(TxtEmpNumber.Text) = 0 Then
   TxtEmpNumber.SetFocus
   Exit Sub
End If
        
'Criteria = "EmpNo = '" & TxtEmpNumber.Text & "' and divcode = '" & thedivcode & "'"
Criteria = "EmpNo = '" & TxtEmpNumber.Text & "'"
Set rsEmpInfoCards = New ADODB.Recordset
Set rsEmpInfoCards = gconDMIS.Execute("Select * from HRMS_EmpInfo Where " & Criteria)
'rsEmpInfoCards.Find Criteria

If rsEmpInfoCards.EOF Then
   For k = 1 To 150: Beep: Next k
   MsgBox "Employee Number NOT FOUND"
   TxtEmpNumber.Text = ""
   TxtEmpNumber.SetFocus
   Exit Sub
Else
   TxtEmpNumber.Text = rsEmpInfoCards!empno
   TxtEmpName.Text = rsEmpInfoCards!lastname & ", " & rsEmpInfoCards!FIRSTNAME
   percas = rsEmpInfoCards!EMPLEVEL
   If rsEmpInfoCards!ACTIVEINACTIVE = "I" Then
      For k = 1 To 150: Beep: Next k
      MsgBox "Employee NOT ACTIVE", 0
      TxtEmpNumber.SetFocus
      Exit Sub
   End If
   Criteria = "Select * from HRMS_Attend Where right(EmpNo,4) = " & TxtEmpNumber.Text & "AND Month(DateToday) = " & Month(tdate) & "AND Year(DateToday) = " & Year(tdate)
   'Set rsCard = gconDMIS.OpenADODB.Recordset(Criteria)
   Set rsCard = New ADODB.Recordset
   Set rsCard = gconDMIS.Execute(Criteria)
   If Not rsCard.EOF And Not rsCard.BOF Then
   rsCard.MoveFirst
   Do Until rsCard.EOF
      r = Day(rsCard!DateToday)
      If r < 16 Then
         With Grid1
              .Row = r
              .Col = 0
              .Text = Format(rsCard!DateToday, "mm/dd/yy")
              .Col = 1
              .Text = Format(Null2String(rsCard!InAm), "hh:mm AM/PM")
              .Col = 2
              .Text = Format(Null2String(rsCard!OutAm), "hh:mm AM/PM")
              .Col = 3
              .Text = Format(Null2String(rsCard!InPm), "hh:mm AM/PM")
              .Col = 4
              .Text = Format(Null2String(rsCard!OutPM), "hh:mm AM/PM")
         End With
      Else
         With Grid2
              .Row = r - 15
              .Col = 0
              .Text = Format(rsCard!DateToday, "mm/dd/yy")
              .Col = 1
              .Text = Format(Null2String(rsCard!InAm), "hh:mm AM/PM")
              .Col = 2
              .Text = Format(Null2String(rsCard!OutAm), "hh:mm AM/PM")
              .Col = 3
              .Text = Format(Null2String(rsCard!InPm), "hh:mm AM/PM")
              .Col = 4
              .Text = Format(Null2String(rsCard!OutPM), "hh:mm AM/PM")
         End With
      End If
      rsCard.MoveNext
   Loop
   End If
End If
End Sub
