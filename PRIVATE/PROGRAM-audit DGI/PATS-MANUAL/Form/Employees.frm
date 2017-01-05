VERSION 5.00
Begin VB.Form frmEmployees 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employees"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   ForeColor       =   &H8000000F&
   Icon            =   "Employees.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Enabled         =   0   'False
      Height          =   2955
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   6375
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   150
         TabIndex        =   27
         Top             =   240
         Width           =   4815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Status:"
         Height          =   1185
         Left            =   3120
         TabIndex        =   20
         Top             =   1680
         Width           =   3135
         Begin VB.OptionButton Option2 
            Caption         =   "Inactive"
            Height          =   255
            Index           =   1
            Left            =   300
            TabIndex        =   25
            Top             =   570
            Width           =   1935
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Active"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   24
            Top             =   300
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Type of Appointment:"
         Height          =   1185
         Left            =   150
         TabIndex        =   19
         Top             =   1680
         Width           =   2925
         Begin VB.OptionButton Option1 
            Caption         =   "Contractual"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   23
            Top             =   840
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Probationary"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   570
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Regular"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   300
            Width           =   2175
         End
      End
      Begin VB.TextBox txtDesignate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   1320
         Width           =   5115
      End
      Begin VB.TextBox txtEmpname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   990
         Width           =   5115
      End
      Begin VB.TextBox txtEmpno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   5160
         TabIndex        =   26
         Top             =   300
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Designation"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Employee #:"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   690
         Width           =   1335
      End
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   5895
      TabIndex        =   1
      Top             =   3030
      Width           =   5925
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         Height          =   855
         Left            =   5100
         Picture         =   "Employees.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         Height          =   855
         Left            =   4380
         Picture         =   "Employees.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         Height          =   855
         Left            =   3660
         Picture         =   "Employees.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit"
         Height          =   855
         Left            =   2940
         Picture         =   "Employees.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add"
         Height          =   855
         Left            =   2220
         Picture         =   "Employees.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Next"
         Height          =   855
         Left            =   1500
         Picture         =   "Employees.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Find"
         Height          =   855
         Left            =   780
         Picture         =   "Employees.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Prev"
         Height          =   855
         Left            =   60
         Picture         =   "Employees.frx":1850
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   4380
      ScaleHeight     =   945
      ScaleWidth      =   1575
      TabIndex        =   10
      Top             =   3030
      Visible         =   0   'False
      Width           =   1605
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         Height          =   855
         Left            =   780
         Picture         =   "Employees.frx":1B5A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         Height          =   855
         Left            =   60
         Picture         =   "Employees.frx":1E64
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   30
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub StoreMemvars()
If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
   'rsdivref.Find "Divcode = '" & Null2String(rsEmpInfo!divcode) & "'"
   If Not rsdivref.EOF Then
      Combo1.Text = Null2String(rsdivref!division)
      Label5.Caption = Null2String(rsdivref!divcode)
   Else
      Combo1.Text = "No Match"
   End If
   txtEmpno.Text = Null2String(rsEmpInfo!empno)
   TxtEmpName.Text = Null2String(rsEmpInfo!EmpName)
   txtDesignate.Text = Null2String(rsEmpInfo!designate)
   'If Null2String(rsEmpInfo!casper) = "P" Then
   '   Option1(0).Value = True
   'ElseIf Null2String(rsEmpInfo!casper) = "C" Then
   '   Option1(1).Value = True
   'ElseIf Null2String(rsEmpInfo!casper) = "N" Then
   '   Option1(2).Value = True
   'Else
   '   Option1(0).Value = False
   '   Option1(1).Value = False
   '   Option1(2).Value = False
   'End If
  
   If Null2String(rsEmpInfo!ACTIVEINACTIVE) = "A" Then
      Option2(0).Value = True
   ElseIf Null2String(rsEmpInfo!ACTIVEINACTIVE) = "I" Then
      Option2(1).Value = True
   End If
Else
   Frame1.Caption = "New"
   Frame1.Enabled = True
   InitMemvars
   Pic1.Visible = False
   Pic2.Visible = True
End If
End Sub

Sub InitMemvars()
txtEmpno.Text = ""
TxtEmpName.Text = ""
txtDesignate.Text = ""
Option1(0).Value = True
Option2(0).Value = True
End Sub

Function FindDup()
On Error GoTo BFoundErr
Dim OldEmpNo As Integer
OldEmpNo = rsEmpInfo!empno

If Not IsNull(txtEmpno.Text) Then
   'find the code
   rsEmpInfo.Find "empno = " & N2Str2Null(txtEmpno.Text)
   If Not rsEmpInfo.EOF Then
      'found it
      MsgBox "Employee Number already Found.", , "Error encoding"
      FindDup = True
   Else
      FindDup = False
   End If
End If
rsEmpInfo.Find "empno = " & N2Str2Null(OldEmpNo)
Exit Function

BFoundErr:
MsgBox Err.Description
End Function


Private Sub cmdAdd_Click()
Frame1.Caption = "New"
Frame1.Enabled = True
InitMemvars
Pic1.Visible = False
Pic2.Visible = True
txtEmpno.SetFocus
End Sub

Private Sub cmdCancel_Click()
Frame1.Caption = "Info"
Frame1.Enabled = False
Pic1.Visible = True
Pic2.Visible = False
txtEmpno.Enabled = True
StoreMemvars
End Sub

Private Sub cmdDelete_Click()

If MsgBox("Are you sure?", vbYesNo + vbExclamation, "Warning!") = vbYes Then
   gconDMIS.Execute "Delete from HRMS_EmpInfo where empno = " & N2Str2Null(txtEmpno.Text)
   'rsRefresh
   rsEmpInfo.MoveFirst
   StoreMemvars
End If
End Sub

Private Sub cmdEdit_Click()
Frame1.Enabled = True
txtEmpno.Enabled = False
Pic1.Visible = False
Pic2.Visible = True
'txtEmpno.SetFocus
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
On Error GoTo BFoundErr
Dim result As Boolean
Dim Origempno  As Integer
Dim findStr As String

findStr = InputBox("Please enter Name to find...", "Find")
If findStr <> "" Then
   result = False
   If Not IsNull(findStr) Then
      Origempno = N2Str2Null(rsEmpInfo!empno)
      If IsNumeric(findStr) Then
         'find the id
         rsEmpInfo.Find "empno = '" & findStr & "'"
      Else
         'find the name
         rsEmpInfo.Find "empname like '" & findStr & "*'"
      End If
      If rsEmpInfo.EOF Then
         MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
         rsEmpInfo.Find "empno = " & Origempno
      Else
         StoreMemvars
      End If
   End If
End If

Exit Sub
BFoundErr:
   MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
rsEmpInfo.MoveNext
If rsEmpInfo.EOF Then
   rsEmpInfo.MoveLast
   MsgBox "Last record.", vbOKOnly, "Warning"
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsEmpInfo.MovePrevious
If rsEmpInfo.BOF Then
   rsEmpInfo.MoveFirst
   MsgBox "First record.", vbOKOnly, "Warning"
End If
StoreMemvars
End Sub

Private Sub cmdSave_Click()
If Combo1.Text = "" Then
   MsgBox "Division must have a value.", , "Warning"
   Combo1.SetFocus
   Exit Sub
End If

If Not ComboFound() Then
   MsgBox "Division not found.", , "Warning"
   Combo1.SetFocus
   Exit Sub
End If

If txtEmpno.Text = "" Then
   MsgBox "Employee No. must have a value", , "Warning"
   txtEmpno.SetFocus
   Exit Sub
End If

If Frame1.Caption = "New" Then
   If FindDup Then
      txtEmpno.SetFocus
      Exit Sub
   End If
End If

If TxtEmpName.Text = "" Then
   MsgBox "Employee Name must have a value", , "Warning"
   TxtEmpName.SetFocus
   Exit Sub
End If

Dim opt1 As String
Dim opt2 As String
TxtEmpName.Text = N2Str2Null(TxtEmpName.Text)
txtDesignate.Text = N2Str2Null(txtDesignate.Text)
If Option1(0).Value = True Then
   opt1 = "'P'"
ElseIf Option1(1).Value = True Then
   opt1 = "'C'"
ElseIf Option1(2).Value = True Then
   opt1 = "'N'"
Else
   opt1 = "''"
End If
If Option2(0).Value = True Then
   opt2 = "'A'"
ElseIf Option2(1).Value = True Then
   opt2 = "'I'"
Else
   opt2 = "''"
End If

If Frame1.Caption = "New" Then
   gconDMIS.Execute "Insert into HRMS_EmpInfo " & _
                    "(empno, divcode, empname, designate, casper, status)" & _
                    " values (" & N2Str2Null(txtEmpno.Text) & ", " & Label5.Caption & ", " & TxtEmpName.Text & ", " & txtDesignate.Text & " " & _
                    ", " & opt1 & ", " & opt2 & ")"
Else
   gconDMIS.Execute "update HRMS_EmpInfo set " & _
                    "divcode = " & Label5.Caption & ", " & _
                    "empname = " & TxtEmpName.Text & ", " & _
                    "designate = " & txtDesignate.Text & ", " & _
                    "casper = " & opt1 & ", " & _
                    "status = " & opt2 & " " & _
                    "where empno = " & N2Str2Null(txtEmpno.Text)
End If
rsEmpInfo.Find "empno = " & N2Str2Null(txtEmpno.Text)
CmdCancel.Value = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = """" Or Chr(KeyAscii) = "'" Then
   KeyAscii = 0
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Mid(ActiveControl.Name, 1, 3) = "txt" Or Mid(ActiveControl.Name, 1, 3) = "cbo" Then
   If KeyCode = 13 Then
      SendKeys "{TAB}"
   End If
End If
End Sub

Private Sub Form_Load()
CenterMe Me, Me, 0
rsEmpInfo.MoveFirst
FillcboDivref
StoreMemvars
End Sub

Sub FillcboDivref()
Combo1.Clear
rsdivref.MoveFirst
Do While Not rsdivref.EOF
   Combo1.AddItem rsdivref!division
   rsdivref.MoveNext
Loop
End Sub

Function ComboFound()
ComboFound = True
rsdivref.Find "Division = '" & Trim(Combo1.Text) & "'"
If rsdivref.EOF Then
   ComboFound = False
Else
   Label5.Caption = rsdivref!divcode
End If
End Function
